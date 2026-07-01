"""
Telegram-бот для управления спикерами мероприятий.

Один файл для MVP/первого продакшена:
- aiogram 3.x: команды, inline-кнопки, FSM-мастера;
- SQLAlchemy async: SQLite по умолчанию, PostgreSQL через DATABASE_URL;
- APScheduler: напоминания о выступлениях, документах и цветах;
- OpenAI API: распознавание расписаний и обычных текстовых команд;
- загрузка расписаний: xlsx/xls/csv/txt/docx/pdf/jpg/png, с OCR при наличии pytesseract.

Переменные окружения в .env:
BOT_TOKEN=123:...
OPENAI_API_KEY=sk-...
DATABASE_URL=sqlite+aiosqlite:///speaker_bot.db
ALLOWED_USER_IDS=123456789,987654321
OPENAI_MODEL=gpt-4.1-mini
DEFAULT_TIMEZONE=Europe/Moscow
"""

from __future__ import annotations

import asyncio
import json
import logging
import os
import re
import shutil
import tempfile
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from pathlib import Path
from typing import Any, Iterable
from zoneinfo import ZoneInfo

from aiogram import Bot, Dispatcher, F, Router
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.filters import CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import CallbackQuery, Document, Message, PhotoSize
from aiogram.utils.keyboard import InlineKeyboardBuilder
from apscheduler.schedulers.asyncio import AsyncIOScheduler
from dotenv import load_dotenv
from openai import AsyncOpenAI
from sqlalchemy import (
    JSON,
    Boolean,
    Date,
    DateTime,
    Float,
    ForeignKey,
    Integer,
    String,
    Text,
    Time,
    func,
    select,
)
from sqlalchemy.ext.asyncio import AsyncSession, async_sessionmaker, create_async_engine
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column, relationship, selectinload

try:
    import pandas as pd
except Exception:
    pd = None

try:
    import docx
except Exception:
    docx = None

try:
    import pdfplumber
except Exception:
    pdfplumber = None

try:
    import pytesseract
    from PIL import Image
except Exception:
    pytesseract = None
    Image = None


load_dotenv()
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("speaker-event-bot")

BOT_TOKEN = os.getenv("BOT_TOKEN", "").strip()
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "").strip()
OPENAI_MODEL = os.getenv("OPENAI_MODEL", "gpt-4.1-mini").strip()
DEFAULT_TZ = os.getenv("DEFAULT_TIMEZONE", "Europe/Moscow").strip()
ALLOWED_USER_IDS = {
    int(x.strip())
    for x in os.getenv("ALLOWED_USER_IDS", "").split(",")
    if x.strip().isdigit()
}

DATABASE_URL = os.getenv("DATABASE_URL", "sqlite+aiosqlite:///speaker_bot.db").strip()
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql+asyncpg://", 1)
elif DATABASE_URL.startswith("postgresql://"):
    DATABASE_URL = DATABASE_URL.replace("postgresql://", "postgresql+asyncpg://", 1)
elif DATABASE_URL.startswith("sqlite:///"):
    DATABASE_URL = DATABASE_URL.replace("sqlite:///", "sqlite+aiosqlite:///", 1)

UPLOAD_DIR = Path(os.getenv("UPLOAD_DIR", "uploads"))
UPLOAD_DIR.mkdir(exist_ok=True)

DOC_FIELDS = [
    ("contract_act_approved", "Договор / акт согласованы"),
    ("contract_act_signed", "Договор / акт подписаны"),
    ("invoice_or_receipt_received", "Счет / чек получен"),
    ("scans_sent", "Сканы в СЭД / по почте"),
    ("paid", "Оплачен"),
    ("originals_received", "ОРИГИНАЛЫ НА РУКАХ?"),
    ("originals_to_accounting", "Оригиналы в бухгалтерию"),
]
CONTRACT_TYPES = ["ИП", "ИП с НДС", "Самозанятый", "ГПХ"]
SETUP_HINT = "презентация, микрофон, петличка, кликер, барный стул, вода, цветы, другое"
PENDING_AI_ACTIONS: dict[int, dict[str, Any]] = {}


class Base(DeclarativeBase):
    pass


class Module(Base):
    __tablename__ = "modules"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    name: Mapped[str] = mapped_column(String(255), index=True)
    start_date: Mapped[date | None] = mapped_column(Date, nullable=True)
    end_date: Mapped[date | None] = mapped_column(Date, nullable=True)
    dates_label: Mapped[str] = mapped_column(String(255), default="")
    location: Mapped[str] = mapped_column(String(255), default="")
    description: Mapped[str] = mapped_column(Text, default="")
    owner_user_id: Mapped[int] = mapped_column(Integer, index=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)
    updated_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    module_speakers: Mapped[list["ModuleSpeaker"]] = relationship(back_populates="module", cascade="all, delete-orphan")
    schedule_items: Mapped[list["ScheduleItem"]] = relationship(back_populates="module", cascade="all, delete-orphan")


class Speaker(Base):
    __tablename__ = "speakers"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    full_name: Mapped[str] = mapped_column(String(255), index=True)
    gender: Mapped[str] = mapped_column(String(32), default="")
    phone: Mapped[str] = mapped_column(String(64), default="")
    telegram: Mapped[str] = mapped_column(String(128), default="")
    email: Mapped[str] = mapped_column(String(255), default="")
    organization: Mapped[str] = mapped_column(String(255), default="")
    position: Mapped[str] = mapped_column(String(255), default="")
    comment: Mapped[str] = mapped_column(Text, default="")
    owner_user_id: Mapped[int] = mapped_column(Integer, index=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)
    updated_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    module_links: Mapped[list["ModuleSpeaker"]] = relationship(back_populates="speaker", cascade="all, delete-orphan")


class ModuleSpeaker(Base):
    __tablename__ = "module_speakers"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    module_id: Mapped[int] = mapped_column(ForeignKey("modules.id", ondelete="CASCADE"), index=True)
    speaker_id: Mapped[int] = mapped_column(ForeignKey("speakers.id", ondelete="CASCADE"), index=True)
    topic: Mapped[str] = mapped_column(String(500), default="")
    is_paid: Mapped[bool] = mapped_column(Boolean, default=False)
    amount: Mapped[float] = mapped_column(Float, default=0)
    contract_type: Mapped[str] = mapped_column(String(64), default="")
    setup: Mapped[list[str]] = mapped_column(JSON, default=list)
    flower_required: Mapped[bool] = mapped_column(Boolean, default=False)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)
    updated_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    module: Mapped[Module] = relationship(back_populates="module_speakers")
    speaker: Mapped[Speaker] = relationship(back_populates="module_links")
    document_status: Mapped["DocumentStatus"] = relationship(back_populates="module_speaker", cascade="all, delete-orphan")


class ScheduleItem(Base):
    __tablename__ = "schedule_items"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    module_id: Mapped[int] = mapped_column(ForeignKey("modules.id", ondelete="CASCADE"), index=True)
    speaker_id: Mapped[int | None] = mapped_column(ForeignKey("speakers.id", ondelete="SET NULL"), nullable=True, index=True)
    date: Mapped[date] = mapped_column(Date, index=True)
    start_time: Mapped[time | None] = mapped_column(Time, nullable=True)
    end_time: Mapped[time | None] = mapped_column(Time, nullable=True)
    title: Mapped[str] = mapped_column(String(500), default="")
    location: Mapped[str] = mapped_column(String(255), default="")
    format: Mapped[str] = mapped_column(String(128), default="")
    comment: Mapped[str] = mapped_column(Text, default="")
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)
    updated_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    module: Mapped[Module] = relationship(back_populates="schedule_items")
    speaker: Mapped[Speaker | None] = relationship()


class DocumentStatus(Base):
    __tablename__ = "document_statuses"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    module_speaker_id: Mapped[int] = mapped_column(ForeignKey("module_speakers.id", ondelete="CASCADE"), unique=True)
    contract_act_approved: Mapped[bool] = mapped_column(Boolean, default=False)
    contract_act_signed: Mapped[bool] = mapped_column(Boolean, default=False)
    invoice_or_receipt_received: Mapped[bool] = mapped_column(Boolean, default=False)
    scans_sent: Mapped[bool] = mapped_column(Boolean, default=False)
    paid: Mapped[bool] = mapped_column(Boolean, default=False)
    originals_received: Mapped[bool] = mapped_column(Boolean, default=False)
    originals_to_accounting: Mapped[bool] = mapped_column(Boolean, default=False)
    progress_percent: Mapped[int] = mapped_column(Integer, default=0)
    updated_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    module_speaker: Mapped[ModuleSpeaker] = relationship(back_populates="document_status")


class Reminder(Base):
    __tablename__ = "reminders"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    module_id: Mapped[int] = mapped_column(ForeignKey("modules.id", ondelete="CASCADE"), index=True)
    speaker_id: Mapped[int | None] = mapped_column(ForeignKey("speakers.id", ondelete="SET NULL"), nullable=True, index=True)
    schedule_item_id: Mapped[int | None] = mapped_column(ForeignKey("schedule_items.id", ondelete="CASCADE"), nullable=True)
    reminder_type: Mapped[str] = mapped_column(String(64), index=True)
    remind_at: Mapped[datetime] = mapped_column(DateTime, index=True)
    is_sent: Mapped[bool] = mapped_column(Boolean, default=False)
    target_user_id: Mapped[int] = mapped_column(Integer, index=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)
    sent_at: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)


class UploadedFile(Base):
    __tablename__ = "uploaded_files"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    module_id: Mapped[int] = mapped_column(ForeignKey("modules.id", ondelete="CASCADE"), index=True)
    file_name: Mapped[str] = mapped_column(String(255))
    file_type: Mapped[str] = mapped_column(String(32))
    file_path: Mapped[str] = mapped_column(String(500))
    parsed_json: Mapped[dict[str, Any]] = mapped_column(JSON, default=dict)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)


class ActionLog(Base):
    __tablename__ = "action_logs"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    user_id: Mapped[int] = mapped_column(Integer, index=True)
    action_type: Mapped[str] = mapped_column(String(100))
    entity_type: Mapped[str] = mapped_column(String(100))
    entity_id: Mapped[int | None] = mapped_column(Integer, nullable=True)
    old_value: Mapped[dict[str, Any] | None] = mapped_column(JSON, nullable=True)
    new_value: Mapped[dict[str, Any] | None] = mapped_column(JSON, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)


class AppSetting(Base):
    __tablename__ = "settings"

    user_id: Mapped[int] = mapped_column(Integer, primary_key=True)
    timezone: Mapped[str] = mapped_column(String(64), default=DEFAULT_TZ)
    performance_reminder_minutes: Mapped[int] = mapped_column(Integer, default=40)
    docs_reminder_minutes: Mapped[int] = mapped_column(Integer, default=30)
    flowers_reminder_minutes: Mapped[int] = mapped_column(Integer, default=30)
    ai_enabled: Mapped[bool] = mapped_column(Boolean, default=True)
    post_module_reminders_enabled: Mapped[bool] = mapped_column(Boolean, default=True)
    post_module_interval_days: Mapped[int] = mapped_column(Integer, default=2)


engine = create_async_engine(DATABASE_URL, echo=False)
SessionLocal = async_sessionmaker(engine, expire_on_commit=False)
router = Router()
openai_client = AsyncOpenAI(api_key=OPENAI_API_KEY) if OPENAI_API_KEY else None


class ModuleCreate(StatesGroup):
    name = State()
    dates = State()
    location = State()
    description = State()


class ModuleEdit(StatesGroup):
    field = State()


class SpeakerCreate(StatesGroup):
    full_name = State()
    gender = State()
    phone = State()
    telegram = State()
    email = State()
    organization = State()
    position = State()
    topic = State()
    is_paid = State()
    amount = State()
    contract_type = State()
    setup = State()
    comment = State()


class SpeakerEdit(StatesGroup):
    setup = State()
    amount = State()
    data = State()


class ScheduleUpload(StatesGroup):
    waiting_file = State()


class ManualSchedule(StatesGroup):
    item_date = State()
    start_time = State()
    end_time = State()
    title = State()
    speaker_name = State()
    location = State()
    item_format = State()
    comment = State()


class SettingsEdit(StatesGroup):
    timezone = State()
    performance_minutes = State()
    docs_minutes = State()
    flowers_minutes = State()
    post_interval = State()


def html_escape(text: Any) -> str:
    s = "" if text is None else str(text)
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def money(value: float | int | None) -> str:
    if not value:
        return "0 руб."
    return f"{int(value):,}".replace(",", " ") + " руб."


def parse_time(value: str) -> time | None:
    value = (value or "").strip()
    m = re.search(r"(\d{1,2})[:.](\d{2})", value)
    if not m:
        return None
    h, minute = int(m.group(1)), int(m.group(2))
    if 0 <= h <= 23 and 0 <= minute <= 59:
        return time(h, minute)
    return None


def parse_date_value(value: str, fallback_year: int | None = None) -> date | None:
    value = (value or "").strip().lower()
    fallback_year = fallback_year or datetime.now().year
    months = {
        "январ": 1, "феврал": 2, "март": 3, "апрел": 4, "ма": 5, "июн": 6,
        "июл": 7, "август": 8, "сентябр": 9, "октябр": 10, "ноябр": 11, "декабр": 12,
    }
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d.%m.%y", "%d.%m"):
        try:
            parsed = datetime.strptime(value, fmt)
            return date(fallback_year if fmt == "%d.%m" else parsed.year, parsed.month, parsed.day)
        except ValueError:
            pass
    m = re.search(r"(\d{1,2})\s+([а-яё]+)(?:\s+(\d{4}))?", value)
    if m:
        month = next((num for key, num in months.items() if m.group(2).startswith(key)), None)
        if month:
            return date(int(m.group(3) or fallback_year), month, int(m.group(1)))
    return None


def parse_date_range(text: str) -> tuple[date | None, date | None, str]:
    label = text.strip()
    year = datetime.now().year
    dates = re.findall(r"\d{4}-\d{2}-\d{2}|\d{1,2}\.\d{1,2}(?:\.\d{2,4})?|\d{1,2}\s+[а-яё]+(?:\s+\d{4})?", label.lower())
    if len(dates) >= 2:
        return parse_date_value(dates[0], year), parse_date_value(dates[1], year), label
    first = parse_date_value(label, year)
    return first, first, label


def split_setup(text: str | Iterable[str] | None) -> list[str]:
    if not text:
        return []
    if isinstance(text, list):
        raw = text
    else:
        raw = re.split(r"[,;\n]+", str(text))
    return [x.strip().lower() for x in raw if x and x.strip()]


def doc_progress(ds: DocumentStatus | None) -> int:
    if not ds:
        return 0
    done = sum(1 for field, _ in DOC_FIELDS if getattr(ds, field))
    ds.progress_percent = round(done / len(DOC_FIELDS) * 100)
    return ds.progress_percent


async def get_settings(session: AsyncSession, user_id: int) -> AppSetting:
    settings = await session.get(AppSetting, user_id)
    if not settings:
        settings = AppSetting(user_id=user_id)
        session.add(settings)
        await session.flush()
    return settings


async def log_action(
    session: AsyncSession,
    user_id: int,
    action_type: str,
    entity_type: str,
    entity_id: int | None,
    old: dict[str, Any] | None = None,
    new: dict[str, Any] | None = None,
) -> None:
    session.add(ActionLog(user_id=user_id, action_type=action_type, entity_type=entity_type, entity_id=entity_id, old_value=old, new_value=new))


def kb_main() -> InlineKeyboardBuilder:
    kb = InlineKeyboardBuilder()
    items = [
        ("📁 Мои модули", "modules"),
        ("➕ Создать модуль", "module_create"),
        ("👤 Спикеры", "speakers"),
        ("📅 Расписание", "schedule_all"),
        ("💰 Оплаты и документы", "payments_all"),
        ("🌸 Цветы", "flowers_all"),
        ("🔔 Напоминания", "reminders"),
        ("⚙️ Настройки", "settings"),
    ]
    for text, data in items:
        kb.button(text=text, callback_data=data)
    kb.adjust(2)
    return kb


def kb_back(module_id: int | None = None) -> InlineKeyboardBuilder:
    kb = InlineKeyboardBuilder()
    if module_id:
        kb.button(text="⬅️ Назад", callback_data=f"module:{module_id}")
    kb.button(text="🏠 Главное меню", callback_data="menu")
    kb.adjust(1)
    return kb


def kb_module(module_id: int) -> InlineKeyboardBuilder:
    kb = InlineKeyboardBuilder()
    buttons = [
        ("👤 Спикеры модуля", f"module_speakers:{module_id}"),
        ("📅 Расписание по дням", f"schedule_days:{module_id}"),
        ("💰 Оплаты и документы", f"module_payments:{module_id}"),
        ("🌸 Цветы", f"module_flowers:{module_id}"),
        ("➕ Добавить спикера", f"speaker_create:{module_id}"),
        ("📎 Загрузить сетку", f"upload_schedule:{module_id}"),
        ("➕ Добавить слот вручную", f"manual_schedule:{module_id}"),
        ("✏️ Изменить модуль", f"module_edit:{module_id}"),
        ("🗑 Удалить модуль", f"module_delete_ask:{module_id}"),
        ("🏠 Главное меню", "menu"),
    ]
    for text, data in buttons:
        kb.button(text=text, callback_data=data)
    kb.adjust(2, 2, 2, 1, 2, 1)
    return kb


def kb_speaker(ms_id: int, module_id: int) -> InlineKeyboardBuilder:
    kb = InlineKeyboardBuilder()
    buttons = [
        ("✏️ Изменить данные", f"speaker_edit_data:{ms_id}"),
        ("⚙️ Изменить сетап", f"speaker_edit_setup:{ms_id}"),
        ("💰 Изменить сумму", f"speaker_edit_amount:{ms_id}"),
        ("📄 Статусы документов", f"docs:{ms_id}"),
        ("📅 Перейти к расписанию", f"schedule_days:{module_id}"),
        ("🔔 Настроить напоминания", f"recreate_reminders:{ms_id}"),
        ("🗑 Удалить спикера", f"speaker_delete_ask:{ms_id}"),
        ("⬅️ К модулю", f"module:{module_id}"),
        ("🏠 Главное меню", "menu"),
    ]
    for text, data in buttons:
        kb.button(text=text, callback_data=data)
    kb.adjust(2, 2, 1, 1, 1, 2)
    return kb


async def find_speaker(session: AsyncSession, user_id: int, name: str) -> Speaker | None:
    name_norm = name.strip().lower()
    if not name_norm:
        return None
    speakers = (await session.scalars(select(Speaker).where(Speaker.owner_user_id == user_id))).all()
    for sp in speakers:
        if name_norm in sp.full_name.lower() or sp.full_name.lower() in name_norm:
            return sp
    first_token = name_norm.split()[0]
    for sp in speakers:
        if first_token and first_token in sp.full_name.lower():
            return sp
    return None


async def find_module(session: AsyncSession, user_id: int, name: str) -> Module | None:
    name_norm = name.strip().lower()
    modules = (await session.scalars(select(Module).where(Module.owner_user_id == user_id))).all()
    for module in modules:
        if name_norm in module.name.lower() or module.name.lower() in name_norm:
            return module
    return modules[0] if modules and not name_norm else None


async def get_module_speaker(session: AsyncSession, module_id: int, speaker_id: int) -> ModuleSpeaker | None:
    return await session.scalar(
        select(ModuleSpeaker)
        .options(selectinload(ModuleSpeaker.document_status), selectinload(ModuleSpeaker.speaker), selectinload(ModuleSpeaker.module))
        .where(ModuleSpeaker.module_id == module_id, ModuleSpeaker.speaker_id == speaker_id)
    )


async def ensure_module_speaker(
    session: AsyncSession,
    module: Module,
    speaker: Speaker,
    topic: str = "",
    setup: list[str] | None = None,
    is_paid: bool = False,
    amount: float = 0,
    contract_type: str = "",
) -> ModuleSpeaker:
    link = await get_module_speaker(session, module.id, speaker.id)
    if not link:
        link = ModuleSpeaker(
            module_id=module.id,
            speaker_id=speaker.id,
            topic=topic or "",
            setup=setup or [],
            is_paid=is_paid,
            amount=amount or 0,
            contract_type=contract_type or "",
            flower_required=speaker.gender.lower().startswith("жен"),
        )
        session.add(link)
        await session.flush()
    else:
        if topic:
            link.topic = topic
        if setup:
            link.setup = setup
        if is_paid:
            link.is_paid = True
            link.amount = amount or link.amount
            link.contract_type = contract_type or link.contract_type
    if link.is_paid and not link.document_status:
        session.add(DocumentStatus(module_speaker_id=link.id))
    return link


async def module_card(session: AsyncSession, module_id: int) -> tuple[str, InlineKeyboardBuilder]:
    module = await session.get(Module, module_id)
    if not module:
        return "Модуль не найден.", kb_main()
    links = (
        await session.scalars(
            select(ModuleSpeaker)
            .where(ModuleSpeaker.module_id == module_id)
            .options(selectinload(ModuleSpeaker.speaker), selectinload(ModuleSpeaker.document_status))
        )
    ).all()
    paid = [x for x in links if x.is_paid]
    flowers = [x for x in links if x.flower_required or x.speaker.gender.lower().startswith("жен")]
    progresses = [doc_progress(x.document_status) for x in paid if x.document_status]
    avg_progress = round(sum(progresses) / len(progresses)) if progresses else 0
    now = datetime.now()
    upcoming = await session.scalar(
        select(ScheduleItem)
        .where(ScheduleItem.module_id == module_id, ScheduleItem.date >= now.date())
        .order_by(ScheduleItem.date, ScheduleItem.start_time)
        .options(selectinload(ScheduleItem.speaker))
    )
    lines = [
        f"📁 <b>Модуль:</b> {html_escape(module.name)}",
        f"📅 <b>Даты:</b> {html_escape(module.dates_label or date_range_label(module))}",
        f"📍 <b>Локация:</b> {html_escape(module.location or 'не указана')}",
        "",
        f"👤 <b>Спикеров всего:</b> {len(links)}",
        f"💰 <b>Платных спикеров:</b> {len(paid)}",
        f"🌸 <b>Цветы нужны:</b> {len(flowers)} букет(а)",
        f"📄 <b>Документы закрыты:</b> {avg_progress}%",
    ]
    if upcoming:
        speaker_name = upcoming.speaker.full_name if upcoming.speaker else "Спикер не указан"
        ms = await get_module_speaker(session, module.id, upcoming.speaker_id) if upcoming.speaker_id else None
        lines += [
            "",
            "<b>Ближайшее выступление:</b>",
            f"{time_label(upcoming.start_time)} - {html_escape(speaker_name)}",
            f"Тема: {html_escape(upcoming.title or (ms.topic if ms else ''))}",
            f"Сетап: {html_escape(', '.join(ms.setup) if ms and ms.setup else 'не указан')}",
        ]
    return "\n".join(lines), kb_module(module_id)


def date_range_label(module: Module) -> str:
    if module.start_date and module.end_date and module.start_date != module.end_date:
        return f"{module.start_date:%d.%m.%Y} - {module.end_date:%d.%m.%Y}"
    if module.start_date:
        return f"{module.start_date:%d.%m.%Y}"
    return "не указаны"


def time_label(value: time | None) -> str:
    return value.strftime("%H:%M") if value else "--:--"


async def speaker_card(session: AsyncSession, ms_id: int) -> tuple[str, InlineKeyboardBuilder]:
    ms = await session.scalar(
        select(ModuleSpeaker)
        .where(ModuleSpeaker.id == ms_id)
        .options(selectinload(ModuleSpeaker.speaker), selectinload(ModuleSpeaker.module), selectinload(ModuleSpeaker.document_status))
    )
    if not ms:
        return "Спикер не найден.", kb_main()
    item = await session.scalar(
        select(ScheduleItem)
        .where(ScheduleItem.module_id == ms.module_id, ScheduleItem.speaker_id == ms.speaker_id)
        .order_by(ScheduleItem.date, ScheduleItem.start_time)
    )
    progress = doc_progress(ms.document_status)
    lines = [
        f"👤 <b>Спикер:</b> {html_escape(ms.speaker.full_name)}",
        f"📁 <b>Модуль:</b> {html_escape(ms.module.name)}",
        f"📅 <b>Дата выступления:</b> {item.date.strftime('%d.%m.%Y') if item else 'не указана'}",
        f"⏰ <b>Время:</b> {time_label(item.start_time) if item else '--:--'}-{time_label(item.end_time) if item else '--:--'}",
        f"🎤 <b>Тема:</b> {html_escape(ms.topic or (item.title if item else 'не указана'))}",
        "",
        f"📞 <b>Телефон:</b> {html_escape(ms.speaker.phone or 'не указан')}",
        f"💬 <b>Telegram:</b> {html_escape(ms.speaker.telegram or 'не указан')}",
        f"📧 <b>Email:</b> {html_escape(ms.speaker.email or 'не указан')}",
        "",
        f"💰 <b>Статус:</b> {'платный' if ms.is_paid else 'бесплатный'}",
        f"💵 <b>Сумма:</b> {money(ms.amount)}",
        f"📄 <b>Тип оформления:</b> {html_escape(ms.contract_type or 'не указан')}",
        "",
        "⚙️ <b>Сетап:</b>",
    ]
    lines += [f"- {html_escape(x)}" for x in (ms.setup or ["не указан"])]
    lines += ["", f"📄 <b>Документы:</b> {progress}%"]
    return "\n".join(lines), kb_speaker(ms.id, ms.module_id)


async def send_or_edit(event: Message | CallbackQuery, text: str, kb: InlineKeyboardBuilder | None = None) -> None:
    markup = kb.as_markup() if kb else None
    if isinstance(event, CallbackQuery):
        await event.message.edit_text(text, reply_markup=markup)
        await event.answer()
    else:
        await event.answer(text, reply_markup=markup)


async def access_allowed(user_id: int) -> bool:
    return not ALLOWED_USER_IDS or user_id in ALLOWED_USER_IDS


@router.message(CommandStart())
async def start(message: Message, state: FSMContext) -> None:
    await state.clear()
    if not await access_allowed(message.from_user.id):
        await message.answer("⛔️ Доступ к боту закрыт. Добавьте ваш Telegram ID в ALLOWED_USER_IDS.")
        return
    async with SessionLocal() as session:
        await get_settings(session, message.from_user.id)
        await session.commit()
    await message.answer(
        "Готов помогать со спикерами, расписанием, документами, оплатами и напоминаниями.",
        reply_markup=kb_main().as_markup(),
    )


@router.callback_query(F.data == "menu")
async def menu(callback: CallbackQuery, state: FSMContext) -> None:
    await state.clear()
    await send_or_edit(callback, "Главное меню", kb_main())


@router.callback_query(F.data == "module_create")
async def module_create(callback: CallbackQuery, state: FSMContext) -> None:
    await state.set_state(ModuleCreate.name)
    await callback.message.edit_text("➕ Введите название модуля:", reply_markup=kb_back().as_markup())
    await callback.answer()


@router.message(ModuleCreate.name)
async def module_create_name(message: Message, state: FSMContext) -> None:
    await state.update_data(name=message.text.strip())
    await state.set_state(ModuleCreate.dates)
    await message.answer("📅 Введите даты модуля. Например: 24-25 июня или 2026-06-24 - 2026-06-25")


@router.message(ModuleCreate.dates)
async def module_create_dates(message: Message, state: FSMContext) -> None:
    start, end, label = parse_date_range(message.text)
    await state.update_data(start_date=start.isoformat() if start else None, end_date=end.isoformat() if end else None, dates_label=label)
    await state.set_state(ModuleCreate.location)
    await message.answer("📍 Введите локацию:")


@router.message(ModuleCreate.location)
async def module_create_location(message: Message, state: FSMContext) -> None:
    await state.update_data(location=message.text.strip())
    await state.set_state(ModuleCreate.description)
    await message.answer("📝 Добавьте короткое описание или отправьте «-»:")


@router.message(ModuleCreate.description)
async def module_create_finish(message: Message, state: FSMContext) -> None:
    data = await state.get_data()
    async with SessionLocal() as session:
        module = Module(
            name=data["name"],
            start_date=date.fromisoformat(data["start_date"]) if data.get("start_date") else None,
            end_date=date.fromisoformat(data["end_date"]) if data.get("end_date") else None,
            dates_label=data.get("dates_label", ""),
            location=data.get("location", ""),
            description="" if message.text.strip() == "-" else message.text.strip(),
            owner_user_id=message.from_user.id,
        )
        session.add(module)
        await session.flush()
        await log_action(session, message.from_user.id, "create", "module", module.id, new={"name": module.name})
        await session.commit()
        text, kb = await module_card(session, module.id)
    await state.clear()
    await message.answer(text, reply_markup=kb.as_markup())


@router.callback_query(F.data == "modules")
async def modules_list(callback: CallbackQuery) -> None:
    async with SessionLocal() as session:
        modules = (
            await session.scalars(
                select(Module).where(Module.owner_user_id == callback.from_user.id).order_by(Module.start_date, Module.created_at)
            )
        ).all()
    kb = InlineKeyboardBuilder()
    for module in modules:
        kb.button(text=f"{module.name} | {module.dates_label or date_range_label(module)}", callback_data=f"module:{module.id}")
    kb.button(text="➕ Создать модуль", callback_data="module_create")
    kb.button(text="🏠 Главное меню", callback_data="menu")
    kb.adjust(1)
    await send_or_edit(callback, "📁 <b>Мои модули</b>" if modules else "Пока нет модулей.", kb)


@router.callback_query(F.data.startswith("module:"))
async def module_open(callback: CallbackQuery) -> None:
    module_id = int(callback.data.split(":")[1])
    async with SessionLocal() as session:
        text, kb = await module_card(session, module_id)
    await send_or_edit(callback, text, kb)


@router.callback_query(F.data.startswith("module_delete_ask:"))
async def module_delete_ask(callback: CallbackQuery) -> None:
    module_id = int(callback.data.split(":")[1])
    kb = InlineKeyboardBuilder()
    kb.button(text="✅ Да, удалить", callback_data=f"module_delete:{module_id}")
    kb.button(text="❌ Отмена", callback_data=f"module:{module_id}")
    kb.adjust(1)
    await send_or_edit(callback, "Удалить модуль? Это также удалит расписание, привязки, документы и напоминания.", kb)


@router.callback_query(F.data.startswith("module_delete:"))
async def module_delete(callback: CallbackQuery) -> None:
    module_id = int(callback.data.split(":")[1])
    async with SessionLocal() as session:
        module = await session.get(Module, module_id)
        if module and module.owner_user_id == callback.from_user.id:
            await log_action(session, callback.from_user.id, "delete", "module", module_id, old={"name": module.name})
            await session.delete(module)
            await session.commit()
    await modules_list(callback)


@router.callback_query(F.data.startswith("module_edit:"))
async def module_edit(callback: CallbackQuery) -> None:
    module_id = int(callback.data.split(":")[1])
    kb = InlineKeyboardBuilder()
    for field, title in [("name", "Название"), ("dates", "Даты"), ("location", "Локация"), ("description", "Описание")]:
        kb.button(text=title, callback_data=f"module_edit_field:{module_id}:{field}")
    kb.button(text="⬅️ Назад", callback_data=f"module:{module_id}")
    kb.adjust(2, 2, 1)
    await send_or_edit(callback, "Что изменить в модуле?", kb)


@router.callback_query(F.data.startswith("module_edit_field:"))
async def module_edit_field(callback: CallbackQuery, state: FSMContext) -> None:
    _, module_id, field = callback.data.split(":")
    await state.update_data(module_id=int(module_id), field=field)
    await state.set_state(ModuleEdit.field)
    await callback.message.edit_text("Введите новое значение:", reply_markup=kb_back(int(module_id)).as_markup())
    await callback.answer()


@router.message(ModuleEdit.field)
async def module_edit_save(message: Message, state: FSMContext) -> None:
    data = await state.get_data()
    async with SessionLocal() as session:
        module = await session.get(Module, data["module_id"])
        if module:
            old = {"name": module.name, "dates": module.dates_label, "location": module.location, "description": module.description}
            field = data["field"]
            if field == "dates":
                start, end, label = parse_date_range(message.text)
                module.start_date, module.end_date, module.dates_label = start, end, label
            elif field == "name":
                module.name = message.text.strip()
            elif field == "location":
                module.location = message.text.strip()
            else:
                module.description = message.text.strip()
            await log_action(session, message.from_user.id, "update", "module", module.id, old=old, new={field: message.text.strip()})
            await session.commit()
            text, kb = await module_card(session, module.id)
        else:
            text, kb = "Модуль не найден.", kb_main()
    await state.clear()
    await message.answer(text, reply_markup=kb.as_markup())


@router.callback_query(F.data.startswith("speaker_create:"))
async def speaker_create(callback: CallbackQuery, state: FSMContext) -> None:
    module_id = int(callback.data.split(":")[1])
    await state.update_data(module_id=module_id)
    await state.set_state(SpeakerCreate.full_name)
    await callback.message.edit_text("👤 Введите ФИО спикера:", reply_markup=kb_back(module_id).as_markup())
    await callback.answer()


@router.message(SpeakerCreate.full_name)
async def speaker_full_name(message: Message, state: FSMContext) -> None:
    await state.update_data(full_name=message.text.strip())
    await state.set_state(SpeakerCreate.gender)
    kb = InlineKeyboardBuilder()
    kb.button(text="Мужчина", callback_data="gender:мужчина")
    kb.button(text="Женщина", callback_data="gender:женщина")
    kb.adjust(2)
    await message.answer("Пол:", reply_markup=kb.as_markup())


@router.callback_query(SpeakerCreate.gender, F.data.startswith("gender:"))
async def speaker_gender(callback: CallbackQuery, state: FSMContext) -> None:
    await state.update_data(gender=callback.data.split(":")[1])
    await state.set_state(SpeakerCreate.phone)
    await callback.message.edit_text("📞 Телефон:")
    await callback.answer()


@router.message(SpeakerCreate.phone)
async def speaker_phone(message: Message, state: FSMContext) -> None:
    await state.update_data(phone=message.text.strip())
    await state.set_state(SpeakerCreate.telegram)
    await message.answer("💬 Telegram:")


@router.message(SpeakerCreate.telegram)
async def speaker_tg(message: Message, state: FSMContext) -> None:
    await state.update_data(telegram=message.text.strip())
    await state.set_state(SpeakerCreate.email)
    await message.answer("📧 Email:")


@router.message(SpeakerCreate.email)
async def speaker_email(message: Message, state: FSMContext) -> None:
    await state.update_data(email=message.text.strip())
    await state.set_state(SpeakerCreate.organization)
    await message.answer("🏢 Организация:")


@router.message(SpeakerCreate.organization)
async def speaker_org(message: Message, state: FSMContext) -> None:
    await state.update_data(organization=message.text.strip())
    await state.set_state(SpeakerCreate.position)
    await message.answer("💼 Должность:")


@router.message(SpeakerCreate.position)
async def speaker_position(message: Message, state: FSMContext) -> None:
    await state.update_data(position=message.text.strip())
    await state.set_state(SpeakerCreate.topic)
    await message.answer("🎤 Тема выступления:")


@router.message(SpeakerCreate.topic)
async def speaker_topic(message: Message, state: FSMContext) -> None:
    await state.update_data(topic=message.text.strip())
    await state.set_state(SpeakerCreate.is_paid)
    kb = InlineKeyboardBuilder()
    kb.button(text="💰 Платный", callback_data="paid:1")
    kb.button(text="Бесплатный", callback_data="paid:0")
    kb.adjust(2)
    await message.answer("Спикер платный?", reply_markup=kb.as_markup())


@router.callback_query(SpeakerCreate.is_paid, F.data.startswith("paid:"))
async def speaker_paid(callback: CallbackQuery, state: FSMContext) -> None:
    is_paid = callback.data.endswith(":1")
    await state.update_data(is_paid=is_paid)
    if is_paid:
        await state.set_state(SpeakerCreate.amount)
        await callback.message.edit_text("💵 Укажите сумму договора:")
    else:
        await state.update_data(amount=0, contract_type="")
        await state.set_state(SpeakerCreate.setup)
        await callback.message.edit_text(f"⚙️ Укажите сетап через запятую.\nПодсказка: {SETUP_HINT}")
    await callback.answer()


@router.message(SpeakerCreate.amount)
async def speaker_amount(message: Message, state: FSMContext) -> None:
    amount = float(re.sub(r"[^\d.]", "", message.text.replace(",", ".")) or 0)
    await state.update_data(amount=amount)
    await state.set_state(SpeakerCreate.contract_type)
    kb = InlineKeyboardBuilder()
    for ctype in CONTRACT_TYPES:
        kb.button(text=ctype, callback_data=f"contract:{ctype}")
    kb.adjust(2)
    await message.answer("📄 Тип оформления:", reply_markup=kb.as_markup())


@router.callback_query(SpeakerCreate.contract_type, F.data.startswith("contract:"))
async def speaker_contract(callback: CallbackQuery, state: FSMContext) -> None:
    await state.update_data(contract_type=callback.data.split(":", 1)[1])
    await state.set_state(SpeakerCreate.setup)
    await callback.message.edit_text(f"⚙️ Укажите сетап через запятую.\nПодсказка: {SETUP_HINT}")
    await callback.answer()


@router.message(SpeakerCreate.setup)
async def speaker_setup(message: Message, state: FSMContext) -> None:
    await state.update_data(setup=split_setup(message.text))
    await state.set_state(SpeakerCreate.comment)
    await message.answer("Комментарий / особые требования или «-»:")


@router.message(SpeakerCreate.comment)
async def speaker_finish(message: Message, state: FSMContext) -> None:
    data = await state.get_data()
    async with SessionLocal() as session:
        speaker = Speaker(
            full_name=data["full_name"],
            gender=data["gender"],
            phone=data.get("phone", ""),
            telegram=data.get("telegram", ""),
            email=data.get("email", ""),
            organization=data.get("organization", ""),
            position=data.get("position", ""),
            comment="" if message.text.strip() == "-" else message.text.strip(),
            owner_user_id=message.from_user.id,
        )
        session.add(speaker)
        await session.flush()
        module = await session.get(Module, data["module_id"])
        ms = await ensure_module_speaker(
            session,
            module,
            speaker,
            topic=data.get("topic", ""),
            setup=data.get("setup", []),
            is_paid=data.get("is_paid", False),
            amount=data.get("amount", 0),
            contract_type=data.get("contract_type", ""),
        )
        await log_action(session, message.from_user.id, "create", "speaker", speaker.id, new={"full_name": speaker.full_name})
        await session.commit()
        await recreate_reminders_for_module(session, module.id, message.from_user.id)
        await session.commit()
        text, kb = await speaker_card(session, ms.id)
    await state.clear()
    await message.answer(text, reply_markup=kb.as_markup())


@router.callback_query(F.data == "speakers")
async def speakers_all(callback: CallbackQuery) -> None:
    async with SessionLocal() as session:
        links = (
            await session.scalars(
                select(ModuleSpeaker)
                .join(Speaker)
                .join(Module)
                .where(Speaker.owner_user_id == callback.from_user.id)
                .options(selectinload(ModuleSpeaker.speaker), selectinload(ModuleSpeaker.module))
                .order_by(Speaker.full_name)
            )
        ).all()
    kb = InlineKeyboardBuilder()
    for ms in links:
        kb.button(text=f"{ms.speaker.full_name} | {ms.module.name}", callback_data=f"speaker:{ms.id}")
    kb.button(text="🏠 Главное меню", callback_data="menu")
    kb.adjust(1)
    await send_or_edit(callback, "👤 <b>Спикеры</b>" if links else "Спикеров пока нет.", kb)


@router.callback_query(F.data.startswith("module_speakers:"))
async def module_speakers(callback: CallbackQuery) -> None:
    module_id = int(callback.data.split(":")[1])
    async with SessionLocal() as session:
        links = (
            await session.scalars(
                select(ModuleSpeaker)
                .where(ModuleSpeaker.module_id == module_id)
                .options(selectinload(ModuleSpeaker.speaker))
                .order_by(ModuleSpeaker.id)
            )
        ).all()
    kb = InlineKeyboardBuilder()
    for ms in links:
        kb.button(text=ms.speaker.full_name, callback_data=f"speaker:{ms.id}")
    kb.button(text="➕ Добавить спикера", callback_data=f"speaker_create:{module_id}")
    kb.button(text="⬅️ К модулю", callback_data=f"module:{module_id}")
    kb.adjust(1)
    await send_or_edit(callback, "👤 <b>Спикеры модуля</b>" if links else "В модуле пока нет спикеров.", kb)


@router.callback_query(F.data.startswith("speaker:"))
async def speaker_open(callback: CallbackQuery) -> None:
    ms_id = int(callback.data.split(":")[1])
    async with SessionLocal() as session:
        text, kb = await speaker_card(session, ms_id)
    await send_or_edit(callback, text, kb)


@router.callback_query(F.data.startswith("speaker_delete_ask:"))
async def speaker_delete_ask(callback: CallbackQuery) -> None:
    ms_id = int(callback.data.split(":")[1])
    kb = InlineKeyboardBuilder()
    kb.button(text="✅ Да, удалить", callback_data=f"speaker_delete:{ms_id}")
    kb.button(text="❌ Отмена", callback_data=f"speaker:{ms_id}")
    kb.adjust(1)
    await send_or_edit(callback, "Удалить спикера из модуля? Бот запросил подтверждение перед важным удалением.", kb)


@router.callback_query(F.data.startswith("speaker_delete:"))
async def speaker_delete(callback: CallbackQuery) -> None:
    ms_id = int(callback.data.split(":")[1])
    async with SessionLocal() as session:
        ms = await session.scalar(select(ModuleSpeaker).where(ModuleSpeaker.id == ms_id).options(selectinload(ModuleSpeaker.speaker)))
        module_id = ms.module_id if ms else None
        if ms:
            await log_action(session, callback.from_user.id, "delete", "module_speaker", ms_id, old={"speaker": ms.speaker.full_name})
            await session.delete(ms)
            await session.commit()
    if module_id:
        await module_open_fake(callback, module_id)
    else:
        await send_or_edit(callback, "Главное меню", kb_main())


async def module_open_fake(callback: CallbackQuery, module_id: int) -> None:
    async with SessionLocal() as session:
        text, kb = await module_card(session, module_id)
    await send_or_edit(callback, text, kb)


@router.callback_query(F.data.startswith("speaker_edit_setup:"))
async def speaker_edit_setup(callback: CallbackQuery, state: FSMContext) -> None:
    ms_id = int(callback.data.split(":")[1])
    await state.update_data(ms_id=ms_id)
    await state.set_state(SpeakerEdit.setup)
    await callback.message.edit_text(f"Введите новый сетап через запятую.\nПодсказка: {SETUP_HINT}")
    await callback.answer()


@router.message(SpeakerEdit.setup)
async def speaker_edit_setup_save(message: Message, state: FSMContext) -> None:
    data = await state.get_data()
    async with SessionLocal() as session:
        ms = await session.get(ModuleSpeaker, data["ms_id"])
        old = {"setup": ms.setup}
        ms.setup = split_setup(message.text)
        ms.flower_required = "цветы" in ms.setup
        await log_action(session, message.from_user.id, "update_setup", "module_speaker", ms.id, old=old, new={"setup": ms.setup})
        await session.commit()
        text, kb = await speaker_card(session, ms.id)
    await state.clear()
    await message.answer(text, reply_markup=kb.as_markup())


@router.callback_query(F.data.startswith("speaker_edit_amount:"))
async def speaker_edit_amount(callback: CallbackQuery, state: FSMContext) -> None:
    ms_id = int(callback.data.split(":")[1])
    await state.update_data(ms_id=ms_id)
    await state.set_state(SpeakerEdit.amount)
    await callback.message.edit_text("Введите новую сумму:")
    await callback.answer()


@router.message(SpeakerEdit.amount)
async def speaker_edit_amount_save(message: Message, state: FSMContext) -> None:
    data = await state.get_data()
    amount = float(re.sub(r"[^\d.]", "", message.text.replace(",", ".")) or 0)
    async with SessionLocal() as session:
        ms = await session.scalar(select(ModuleSpeaker).where(ModuleSpeaker.id == data["ms_id"]).options(selectinload(ModuleSpeaker.document_status)))
        old = {"amount": ms.amount}
        ms.amount = amount
        ms.is_paid = amount > 0
        if ms.is_paid and not ms.document_status:
            session.add(DocumentStatus(module_speaker_id=ms.id))
        await log_action(session, message.from_user.id, "update_amount", "module_speaker", ms.id, old=old, new={"amount": amount})
        await session.commit()
        text, kb = await speaker_card(session, ms.id)
    await state.clear()
    await message.answer(text, reply_markup=kb.as_markup())


@router.callback_query(F.data.startswith("speaker_edit_data:"))
async def speaker_edit_data(callback: CallbackQuery) -> None:
    ms_id = int(callback.data.split(":")[1])
    kb = InlineKeyboardBuilder()
    kb.button(text="ФИО", callback_data=f"speaker_field:{ms_id}:full_name")
    kb.button(text="Телефон", callback_data=f"speaker_field:{ms_id}:phone")
    kb.button(text="Telegram", callback_data=f"speaker_field:{ms_id}:telegram")
    kb.button(text="Email", callback_data=f"speaker_field:{ms_id}:email")
    kb.button(text="Тема", callback_data=f"speaker_field:{ms_id}:topic")
    kb.button(text="Комментарий", callback_data=f"speaker_field:{ms_id}:comment")
    kb.button(text="⬅️ Назад", callback_data=f"speaker:{ms_id}")
    kb.adjust(2, 2, 2, 1)
    await send_or_edit(callback, "Что изменить?", kb)


@router.callback_query(F.data.startswith("speaker_field:"))
async def speaker_edit_field(callback: CallbackQuery, state: FSMContext) -> None:
    _, ms_id, field = callback.data.split(":")
    await state.update_data(ms_id=int(ms_id), field=field)
    await state.set_state(SpeakerEdit.data)
    await callback.message.edit_text("Введите новое значение:")
    await callback.answer()


@router.message(SpeakerEdit.data)
async def speaker_edit_data_save(message: Message, state: FSMContext) -> None:
    data = await state.get_data()
    async with SessionLocal() as session:
        ms = await session.scalar(select(ModuleSpeaker).where(ModuleSpeaker.id == data["ms_id"]).options(selectinload(ModuleSpeaker.speaker)))
        field = data["field"]
        target = ms if field == "topic" else ms.speaker
        old = {field: getattr(target, field)}
        setattr(target, field, message.text.strip())
        await log_action(session, message.from_user.id, "update", "speaker", ms.speaker_id, old=old, new={field: message.text.strip()})
        await session.commit()
        text, kb = await speaker_card(session, ms.id)
    await state.clear()
    await message.answer(text, reply_markup=kb.as_markup())


@router.callback_query(F.data.startswith("docs:"))
async def docs_open(callback: CallbackQuery) -> None:
    ms_id = int(callback.data.split(":")[1])
    async with SessionLocal() as session:
        ms = await session.scalar(
            select(ModuleSpeaker)
            .where(ModuleSpeaker.id == ms_id)
            .options(selectinload(ModuleSpeaker.speaker), selectinload(ModuleSpeaker.document_status))
        )
        if not ms:
            await send_or_edit(callback, "Спикер не найден.", kb_main())
            return
        if not ms.document_status:
            ms.is_paid = True
            ms.document_status = DocumentStatus(module_speaker_id=ms.id)
            await session.commit()
        progress = doc_progress(ms.document_status)
        await session.commit()
        lines = [f"📄 <b>Документы: {html_escape(ms.speaker.full_name)}</b>", ""]
        for field, title in DOC_FIELDS:
            lines.append(f"{'✅' if getattr(ms.document_status, field) else '❌'} {html_escape(title)}")
        lines.append(f"\n<b>Прогресс:</b> {progress}%")
    kb = InlineKeyboardBuilder()
    for field, title in DOC_FIELDS:
        kb.button(text=f"Переключить: {title}", callback_data=f"doc_toggle:{ms_id}:{field}")
    kb.button(text="⬅️ К спикеру", callback_data=f"speaker:{ms_id}")
    kb.adjust(1)
    await send_or_edit(callback, "\n".join(lines), kb)


@router.callback_query(F.data.startswith("doc_toggle:"))
async def doc_toggle(callback: CallbackQuery) -> None:
    _, ms_id, field = callback.data.split(":")
    async with SessionLocal() as session:
        ms = await session.scalar(select(ModuleSpeaker).where(ModuleSpeaker.id == int(ms_id)).options(selectinload(ModuleSpeaker.document_status)))
        if ms and ms.document_status and field in dict(DOC_FIELDS):
            old = {field: getattr(ms.document_status, field)}
            setattr(ms.document_status, field, not getattr(ms.document_status, field))
            doc_progress(ms.document_status)
            await log_action(session, callback.from_user.id, "toggle_document", "document_status", ms.document_status.id, old=old, new={field: getattr(ms.document_status, field)})
            await session.commit()
    await docs_open(callback)


@router.callback_query(F.data.startswith("module_payments:") | (F.data == "payments_all"))
async def payments(callback: CallbackQuery) -> None:
    module_id = None if callback.data == "payments_all" else int(callback.data.split(":")[1])
    async with SessionLocal() as session:
        stmt = (
            select(ModuleSpeaker)
            .join(Speaker)
            .join(Module)
            .where(Module.owner_user_id == callback.from_user.id, ModuleSpeaker.is_paid == True)
            .options(selectinload(ModuleSpeaker.speaker), selectinload(ModuleSpeaker.module), selectinload(ModuleSpeaker.document_status))
        )
        if module_id:
            stmt = stmt.where(ModuleSpeaker.module_id == module_id)
        links = (await session.scalars(stmt)).all()
    total = sum(x.amount for x in links)
    avg = round(sum(doc_progress(x.document_status) for x in links) / len(links)) if links else 0
    title = "💰 <b>Оплаты и документы</b>"
    if module_id and links:
        title += f"\n📁 Модуль: {html_escape(links[0].module.name)}"
    lines = [title, "", f"Платных спикеров: {len(links)}", f"Общая сумма: {money(total)}", f"Средний прогресс документов: {avg}%", ""]
    kb = InlineKeyboardBuilder()
    for i, ms in enumerate(links, 1):
        lines += [
            f"{i}. <b>{html_escape(ms.speaker.full_name)}</b>",
            f"Сумма: {money(ms.amount)}",
            f"Тип: {html_escape(ms.contract_type or 'не указан')}",
            f"Прогресс: {doc_progress(ms.document_status)}%",
            "",
        ]
        kb.button(text=f"📄 {ms.speaker.full_name}", callback_data=f"docs:{ms.id}")
    kb.button(text="🏠 Главное меню", callback_data="menu")
    kb.adjust(1)
    await send_or_edit(callback, "\n".join(lines) if links else "Платных спикеров пока нет.", kb)


@router.callback_query(F.data.startswith("module_flowers:") | (F.data == "flowers_all"))
async def flowers(callback: CallbackQuery) -> None:
    module_id = None if callback.data == "flowers_all" else int(callback.data.split(":")[1])
    async with SessionLocal() as session:
        stmt = (
            select(ModuleSpeaker)
            .join(Speaker)
            .join(Module)
            .where(Module.owner_user_id == callback.from_user.id)
            .options(selectinload(ModuleSpeaker.speaker), selectinload(ModuleSpeaker.module))
        )
        if module_id:
            stmt = stmt.where(ModuleSpeaker.module_id == module_id)
        links = [x for x in (await session.scalars(stmt)).all() if x.flower_required or x.speaker.gender.lower().startswith("жен")]
        rows = []
        for ms in links:
            item = await session.scalar(
                select(ScheduleItem)
                .where(ScheduleItem.module_id == ms.module_id, ScheduleItem.speaker_id == ms.speaker_id)
                .order_by(ScheduleItem.date, ScheduleItem.start_time)
            )
            rows.append((ms, item))
    module_name = rows[0][0].module.name if module_id and rows else "всем модулям"
    lines = [
        f"🌸 <b>Цветы по модулю: {html_escape(module_name)}</b>",
        "",
        f"Всего женщин-спикеров: {len(rows)}",
        f"Нужно заказать букетов: {len(rows)}",
        "",
        "Список:",
    ]
    for i, (ms, item) in enumerate(rows, 1):
        when = f"{item.date:%d.%m.%Y}, {time_label(item.start_time)}-{time_label(item.end_time)}" if item else "время не указано"
        lines.append(f"{i}. {html_escape(ms.speaker.full_name)} - {when}")
    await send_or_edit(callback, "\n".join(lines), kb_back(module_id))


@router.callback_query(F.data.startswith("upload_schedule:"))
async def upload_schedule(callback: CallbackQuery, state: FSMContext) -> None:
    module_id = int(callback.data.split(":")[1])
    await state.update_data(module_id=module_id)
    await state.set_state(ScheduleUpload.waiting_file)
    await callback.message.edit_text(
        "📎 Загрузите файл расписания: xlsx, xls, csv, txt, docx, pdf, jpg или png.",
        reply_markup=kb_back(module_id).as_markup(),
    )
    await callback.answer()


@router.message(ScheduleUpload.waiting_file, F.document | F.photo)
async def receive_schedule_file(message: Message, state: FSMContext, bot: Bot) -> None:
    data = await state.get_data()
    module_id = data["module_id"]
    src_name = "schedule.jpg"
    if message.document:
        doc: Document = message.document
        src_name = doc.file_name or f"file_{doc.file_id}"
        file = await bot.get_file(doc.file_id)
    else:
        photo: PhotoSize = message.photo[-1]
        src_name = f"photo_{photo.file_id}.jpg"
        file = await bot.get_file(photo.file_id)
    safe_name = re.sub(r"[^A-Za-zА-Яа-я0-9_.-]+", "_", src_name)
    dst = UPLOAD_DIR / f"{datetime.utcnow():%Y%m%d%H%M%S}_{safe_name}"
    await bot.download_file(file.file_path, destination=dst)

    await message.answer("Файл получен. Извлекаю текст и привожу расписание к единому JSON-формату.")
    extracted = extract_file_text(dst)
    async with SessionLocal() as session:
        module = await session.get(Module, module_id)
        parsed = await parse_schedule_with_ai(module, extracted)
        upload = UploadedFile(
            module_id=module_id,
            file_name=src_name,
            file_type=dst.suffix.lower().lstrip("."),
            file_path=str(dst),
            parsed_json=parsed,
        )
        session.add(upload)
        await session.flush()
        await session.commit()
        preview = schedule_preview(parsed)
    kb = InlineKeyboardBuilder()
    kb.button(text="✅ Сохранить расписание", callback_data=f"schedule_confirm:{upload.id}")
    kb.button(text="❌ Отмена", callback_data=f"module:{module_id}")
    kb.adjust(1)
    await state.clear()
    await message.answer(preview, reply_markup=kb.as_markup())


def extract_file_text(path: Path) -> str:
    ext = path.suffix.lower()
    try:
        if ext in {".csv"}:
            return path.read_text(encoding="utf-8", errors="ignore")
        if ext in {".txt"}:
            return path.read_text(encoding="utf-8", errors="ignore")
        if ext in {".xlsx", ".xls"} and pd is not None:
            frames = pd.read_excel(path, sheet_name=None, header=None)
            return "\n\n".join(f"Лист: {name}\n{df.fillna('').to_csv(index=False, header=False)}" for name, df in frames.items())
        if ext == ".docx" and docx is not None:
            document = docx.Document(str(path))
            parts = [p.text for p in document.paragraphs if p.text.strip()]
            for table in document.tables:
                for row in table.rows:
                    parts.append(" | ".join(cell.text for cell in row.cells))
            return "\n".join(parts)
        if ext == ".pdf" and pdfplumber is not None:
            return extract_pdf_text_with_layout(path)
        if ext in {".jpg", ".jpeg", ".png"} and pytesseract is not None and Image is not None:
            return pytesseract.image_to_string(Image.open(path), lang="rus+eng")
    except Exception as exc:
        log.exception("file extract failed: %s", exc)
        return f"Не удалось извлечь файл автоматически: {exc}"
    return "Не удалось извлечь текст: установите нужную библиотеку или OCR. Можно отправить расписание текстом."


def extract_pdf_text_with_layout(path: Path) -> str:
    """Extract PDF text plus a coordinate-based schedule matrix.

    Many event grids are visually clear but text extraction interleaves days:
    one left time column, several date columns, and merged cells that span
    multiple 15-minute rows. This layout-aware representation gives the AI the
    row/column context it needs to recover exact dates and time ranges.
    """
    parts: list[str] = []
    with pdfplumber.open(str(path)) as pdf:
        for page_index, page in enumerate(pdf.pages, 1):
            raw_text = page.extract_text(x_tolerance=2, y_tolerance=3) or ""
            parts.append(f"\n=== PAGE {page_index} RAW_TEXT ===\n{raw_text}")

            words = page.extract_words(
                x_tolerance=1,
                y_tolerance=2,
                keep_blank_chars=False,
                use_text_flow=False,
            )
            matrix = build_schedule_matrix_from_words(words, page.width)
            if matrix:
                parts.append(f"\n=== PAGE {page_index} LAYOUT_MATRIX ===")
                parts.append(matrix)

            tables = page.extract_tables() or []
            if tables:
                parts.append(f"\n=== PAGE {page_index} TABLES ===")
                for table in tables:
                    for row in table:
                        parts.append(" | ".join((cell or "").strip() for cell in row))
    return "\n".join(parts)


def build_schedule_matrix_from_words(words: list[dict[str, Any]], page_width: float) -> str:
    date_re = re.compile(r"\b\d{1,2}\.\d{1,2}\.\d{4}\b")
    time_re = re.compile(r"^\d{1,2}:\d{2}\s*[-–]\s*\d{1,2}:\d{2}$")

    date_headers = []
    for word in words:
        if date_re.fullmatch(word["text"]):
            date_headers.append(
                {
                    "date": word["text"],
                    "x0": float(word["x0"]),
                    "x1": float(word["x1"]),
                    "center": (float(word["x0"]) + float(word["x1"])) / 2,
                    "top": float(word["top"]),
                }
            )
    date_headers = sorted(date_headers, key=lambda x: x["center"])
    if len(date_headers) < 2:
        return ""

    time_words = []
    first_date_x = date_headers[0]["x0"]
    for word in words:
        text = normalize_extracted_time(word["text"])
        if time_re.fullmatch(text) and float(word["x1"]) < first_date_x:
            start, end = [x.strip() for x in re.split(r"[-–]", text, maxsplit=1)]
            time_words.append(
                {
                    "start": start,
                    "end": end,
                    "x1": float(word["x1"]),
                    "top": float(word["top"]),
                    "bottom": float(word["bottom"]),
                    "center": (float(word["top"]) + float(word["bottom"])) / 2,
                }
            )
    time_words = sorted(time_words, key=lambda x: x["center"])
    if len(time_words) < 2:
        return ""

    time_col_right = max(t["x1"] for t in time_words)
    column_ranges = []
    for i, header in enumerate(date_headers):
        left = (
            time_col_right + 2
            if i == 0
            else (date_headers[i - 1]["center"] + header["center"]) / 2
        )
        right = (
            page_width
            if i == len(date_headers) - 1
            else (header["center"] + date_headers[i + 1]["center"]) / 2
        )
        column_ranges.append((header["date"], left, right))

    row_ranges = []
    for i, slot in enumerate(time_words):
        top = (
            slot["top"] - 2
            if i == 0
            else (time_words[i - 1]["center"] + slot["center"]) / 2
        )
        bottom = (
            slot["bottom"] + 2
            if i == len(time_words) - 1
            else (slot["center"] + time_words[i + 1]["center"]) / 2
        )
        row_ranges.append((slot["start"], slot["end"], top, bottom))

    lines = [
        "The grid below was reconstructed from PDF coordinates.",
        "Columns are dates. Rows are time slots. Empty cells may mean the previous event continues.",
        "If adjacent rows in the same date column contain parts of one visual block, merge them into one schedule item with the first start_time and the last end_time.",
        "",
        "DATES: " + " | ".join(d for d, _, _ in column_ranges),
    ]

    for start, end, top, bottom in row_ranges:
        cells = []
        for day, left, right in column_ranges:
            cell_words = [
                word
                for word in words
                if top <= float(word["bottom"]) < bottom
                and left <= ((float(word["x0"]) + float(word["x1"])) / 2) < right
                and not date_re.fullmatch(word["text"])
            ]
            text = join_positioned_words(cell_words)
            cells.append(f"{day}: {text}" if text else f"{day}:")
        lines.append(f"{start}-{end} || " + " || ".join(cells))

    return "\n".join(lines)


def normalize_extracted_time(text: str) -> str:
    return (
        text.strip()
        .replace("–", "-")
        .replace("—", "-")
        .replace(".", ":")
        .replace(" ", "")
    )


def join_positioned_words(words: list[dict[str, Any]]) -> str:
    if not words:
        return ""
    lines: dict[int, list[dict[str, Any]]] = defaultdict(list)
    for word in words:
        lines[round(float(word["top"]) / 3)].append(word)

    chunks = []
    for _, line_words in sorted(lines.items()):
        row = sorted(line_words, key=lambda w: float(w["x0"]))
        chunks.append(" ".join(w["text"] for w in row))
    text = " / ".join(chunk for chunk in chunks if chunk.strip())
    return re.sub(r"\s+", " ", text).strip()


async def parse_schedule_with_ai(module: Module, text: str) -> dict[str, Any]:
    if openai_client and OPENAI_API_KEY:
        schema = {
            "type": "object",
            "additionalProperties": False,
            "properties": {
                "module_name": {"type": "string"},
                "days": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "additionalProperties": False,
                        "properties": {
                            "date": {"type": "string"},
                            "items": {
                                "type": "array",
                                "items": {
                                    "type": "object",
                                    "additionalProperties": False,
                                    "properties": {
                                        "start_time": {"type": "string"},
                                        "end_time": {"type": "string"},
                                        "title": {"type": "string"},
                                        "speaker_name": {"type": "string"},
                                        "location": {"type": "string"},
                                        "format": {"type": "string"},
                                        "comment": {"type": "string"},
                                    },
                                    "required": ["start_time", "end_time", "title", "speaker_name", "location", "format", "comment"],
                                },
                            },
                        },
                        "required": ["date", "items"],
                    },
                },
            },
            "required": ["module_name", "days"],
        }
        try:
            response = await openai_client.responses.create(
                model=OPENAI_MODEL,
                input=[
                    {
                        "role": "system",
                        "content": (
                            "Ты профессионально извлекаешь расписания мероприятий из PDF/Excel/табличных сеток. "
                            "Верни только валидный JSON по схеме. Даты приводи к YYYY-MM-DD. "
                            "Если в тексте есть LAYOUT_MATRIX, считай ее главным источником: колонки - это даты, строки - временные слоты. "
                            "Текст в одной колонке даты относится именно к этой дате. "
                            "Обязательно восстанавливай длительность событий по соседним строкам одной колонки: "
                            "если название, тема или спикер визуально продолжаются в следующих 15-минутных строках, "
                            "объедини их в один item с start_time первой строки и end_time последней строки. "
                            "Не дроби одно событие на куски по 15 минутам. "
                            "Например, если в 10:00-10:15 указано 'Знакомство', а в 10:15-10:30 '(или мосты)', "
                            "создай один item 10:00-10:30 с title 'Знакомство (или мосты)'. "
                            "Если в строке 7:45-8:00 в колонке 24.07 стоит 'Пробежка с Владимиром Волошиным', "
                            "создай item на 2026-07-24 с start_time 07:45 и end_time 08:00. "
                            "Сохраняй имена спикеров, модераторов, места, формат и комментарии, если они видны. "
                            "Не выдумывай недостающие данные: если спикер не указан явно, оставь speaker_name пустым."
                        ),
                    },
                    {"role": "user", "content": f"Модуль: {module.name}\nДаты: {module.dates_label}\nТекст файла:\n{text[:50000]}"},
                ],
                text={"format": {"type": "json_schema", "name": "schedule", "schema": schema, "strict": True}},
            )
            return json.loads(response.output_text)
        except Exception as exc:
            log.exception("OpenAI schedule parse failed: %s", exc)
    return heuristic_schedule(module, text)


def heuristic_schedule(module: Module, text: str) -> dict[str, Any]:
    layout = heuristic_schedule_from_layout_matrix(module, text)
    if layout.get("days"):
        return layout

    days: dict[str, list[dict[str, str]]] = {}
    current_date = module.start_date or datetime.now().date()
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        found_date = parse_date_value(line)
        if found_date:
            current_date = found_date
        tm = re.search(r"(\d{1,2}[:.]\d{2})\s*[-–]\s*(\d{1,2}[:.]\d{2})", line)
        if tm:
            rest = line[tm.end():].strip(" -–|")
            parts = [x.strip() for x in re.split(r"\s+\|\s+|;", rest) if x.strip()]
            item = {
                "start_time": tm.group(1).replace(".", ":"),
                "end_time": tm.group(2).replace(".", ":"),
                "title": parts[0] if parts else rest,
                "speaker_name": parts[1] if len(parts) > 1 else "",
                "location": parts[2] if len(parts) > 2 else module.location,
                "format": parts[3] if len(parts) > 3 else "",
                "comment": "",
            }
            days.setdefault(current_date.isoformat(), []).append(item)
    return {"module_name": module.name, "days": [{"date": day, "items": items} for day, items in days.items()]}


def heuristic_schedule_from_layout_matrix(module: Module, text: str) -> dict[str, Any]:
    rows = []
    for line in text.splitlines():
        if " || " not in line:
            continue
        time_part, cells_part = line.split(" || ", 1)
        tm = re.fullmatch(r"(\d{1,2}:\d{2})-(\d{1,2}:\d{2})", time_part.strip())
        if not tm:
            continue
        cells = {}
        for raw_cell in cells_part.split(" || "):
            if ":" not in raw_cell:
                continue
            day_raw, cell_text = raw_cell.split(":", 1)
            day = parse_date_value(day_raw.strip())
            if day:
                cells[day.isoformat()] = cell_text.strip()
        rows.append({"start": tm.group(1), "end": tm.group(2), "cells": cells})

    if not rows:
        return {"module_name": module.name, "days": []}

    by_day: dict[str, list[dict[str, str]]] = defaultdict(list)
    active: dict[str, dict[str, str] | None] = defaultdict(lambda: None)

    for row in rows:
        for day, cell_text in row["cells"].items():
            current = active[day]
            if not cell_text:
                if current:
                    current["end_time"] = row["end"]
                continue

            should_continue = bool(
                current
                and (
                    cell_text.startswith(("(", "-", "и ", "или "))
                    or cell_text[:1].islower()
                    or len(cell_text.split()) <= 3
                )
            )
            if should_continue:
                current["title"] = f"{current['title']} {cell_text}".strip()
                current["end_time"] = row["end"]
            else:
                item = {
                    "start_time": row["start"],
                    "end_time": row["end"],
                    "title": cell_text,
                    "speaker_name": "",
                    "location": module.location or "",
                    "format": "",
                    "comment": "fallback_layout_parse",
                }
                by_day[day].append(item)
                active[day] = item

    return {
        "module_name": module.name,
        "days": [{"date": day, "items": items} for day, items in sorted(by_day.items())],
    }


def schedule_preview(parsed: dict[str, Any]) -> str:
    lines = [f"📅 <b>Распознанное расписание: {html_escape(parsed.get('module_name', ''))}</b>", ""]
    for day in parsed.get("days", []):
        lines.append(f"<b>{html_escape(day.get('date', ''))}</b>")
        for item in day.get("items", []):
            speaker = item.get("speaker_name") or "без спикера"
            lines.append(f"{item.get('start_time', '')}-{item.get('end_time', '')} | {html_escape(speaker)}")
            lines.append(f"Тема: {html_escape(item.get('title', ''))}")
        lines.append("")
    return "\n".join(lines[:80]) or "ИИ не нашел расписание в файле."


@router.callback_query(F.data.startswith("schedule_confirm:"))
async def schedule_confirm(callback: CallbackQuery) -> None:
    upload_id = int(callback.data.split(":")[1])
    async with SessionLocal() as session:
        upload = await session.get(UploadedFile, upload_id)
        if not upload:
            await send_or_edit(callback, "Файл не найден.", kb_main())
            return
        await save_schedule_json(session, upload.module_id, upload.parsed_json, callback.from_user.id)
        await recreate_reminders_for_module(session, upload.module_id, callback.from_user.id)
        await log_action(session, callback.from_user.id, "upload_schedule", "module", upload.module_id, new=upload.parsed_json)
        await session.commit()
        text, kb = await module_card(session, upload.module_id)
    await send_or_edit(callback, "✅ Расписание сохранено.\n\n" + text, kb)


async def save_schedule_json(session: AsyncSession, module_id: int, parsed: dict[str, Any], user_id: int) -> None:
    module = await session.get(Module, module_id)
    for old in (await session.scalars(select(ScheduleItem).where(ScheduleItem.module_id == module_id))).all():
        await session.delete(old)
    await session.flush()
    for day in parsed.get("days", []):
        item_date = parse_date_value(day.get("date", "")) or module.start_date or datetime.now().date()
        for item in day.get("items", []):
            speaker = await find_speaker(session, user_id, item.get("speaker_name", ""))
            if not speaker and item.get("speaker_name"):
                speaker = Speaker(full_name=item["speaker_name"], gender="", owner_user_id=user_id)
                session.add(speaker)
                await session.flush()
            if speaker:
                await ensure_module_speaker(session, module, speaker, topic=item.get("title", ""))
            session.add(
                ScheduleItem(
                    module_id=module_id,
                    speaker_id=speaker.id if speaker else None,
                    date=item_date,
                    start_time=parse_time(item.get("start_time", "")),
                    end_time=parse_time(item.get("end_time", "")),
                    title=item.get("title", ""),
                    location=item.get("location", ""),
                    format=item.get("format", ""),
                    comment=item.get("comment", ""),
                )
            )


@router.callback_query(F.data.startswith("manual_schedule:"))
async def manual_schedule(callback: CallbackQuery, state: FSMContext) -> None:
    module_id = int(callback.data.split(":")[1])
    await state.update_data(module_id=module_id)
    await state.set_state(ManualSchedule.item_date)
    await callback.message.edit_text("📅 Дата слота: например 2026-06-24 или 24 июня")
    await callback.answer()


@router.message(ManualSchedule.item_date)
async def manual_date(message: Message, state: FSMContext) -> None:
    parsed = parse_date_value(message.text)
    if not parsed:
        await message.answer("Не понял дату. Попробуйте формат 2026-06-24.")
        return
    await state.update_data(date=parsed.isoformat())
    await state.set_state(ManualSchedule.start_time)
    await message.answer("⏰ Время начала, например 14:30:")


@router.message(ManualSchedule.start_time)
async def manual_start(message: Message, state: FSMContext) -> None:
    if not parse_time(message.text):
        await message.answer("Не понял время. Например: 14:30")
        return
    await state.update_data(start_time=message.text)
    await state.set_state(ManualSchedule.end_time)
    await message.answer("⏰ Время окончания:")


@router.message(ManualSchedule.end_time)
async def manual_end(message: Message, state: FSMContext) -> None:
    if not parse_time(message.text):
        await message.answer("Не понял время. Например: 15:20")
        return
    await state.update_data(end_time=message.text)
    await state.set_state(ManualSchedule.title)
    await message.answer("🎤 Тема / название слота:")


@router.message(ManualSchedule.title)
async def manual_title(message: Message, state: FSMContext) -> None:
    await state.update_data(title=message.text.strip())
    await state.set_state(ManualSchedule.speaker_name)
    await message.answer("👤 ФИО спикера или «-», если это общий слот:")


@router.message(ManualSchedule.speaker_name)
async def manual_speaker(message: Message, state: FSMContext) -> None:
    await state.update_data(speaker_name="" if message.text.strip() == "-" else message.text.strip())
    await state.set_state(ManualSchedule.location)
    await message.answer("📍 Зал / локация:")


@router.message(ManualSchedule.location)
async def manual_location(message: Message, state: FSMContext) -> None:
    await state.update_data(location=message.text.strip())
    await state.set_state(ManualSchedule.item_format)
    await message.answer("Формат: лекция, панель, мастер-класс или «-»:")


@router.message(ManualSchedule.item_format)
async def manual_format(message: Message, state: FSMContext) -> None:
    await state.update_data(format="" if message.text.strip() == "-" else message.text.strip())
    await state.set_state(ManualSchedule.comment)
    await message.answer("Комментарий или «-»:")


@router.message(ManualSchedule.comment)
async def manual_finish(message: Message, state: FSMContext) -> None:
    data = await state.get_data()
    async with SessionLocal() as session:
        module = await session.get(Module, data["module_id"])
        speaker = await find_speaker(session, message.from_user.id, data.get("speaker_name", ""))
        if not speaker and data.get("speaker_name"):
            speaker = Speaker(full_name=data["speaker_name"], owner_user_id=message.from_user.id)
            session.add(speaker)
            await session.flush()
        if speaker:
            await ensure_module_speaker(session, module, speaker, topic=data["title"])
        item = ScheduleItem(
            module_id=module.id,
            speaker_id=speaker.id if speaker else None,
            date=date.fromisoformat(data["date"]),
            start_time=parse_time(data["start_time"]),
            end_time=parse_time(data["end_time"]),
            title=data["title"],
            location=data["location"],
            format=data["format"],
            comment="" if message.text.strip() == "-" else message.text.strip(),
        )
        session.add(item)
        await session.flush()
        await recreate_reminders_for_module(session, module.id, message.from_user.id)
        await log_action(session, message.from_user.id, "create", "schedule_item", item.id, new={"title": item.title})
        await session.commit()
        text, kb = await module_card(session, module.id)
    await state.clear()
    await message.answer("✅ Слот добавлен.\n\n" + text, reply_markup=kb.as_markup())


@router.callback_query(F.data.startswith("schedule_days:") | (F.data == "schedule_all"))
async def schedule_days(callback: CallbackQuery) -> None:
    module_id = None if callback.data == "schedule_all" else int(callback.data.split(":")[1])
    async with SessionLocal() as session:
        stmt = select(ScheduleItem).join(Module).where(Module.owner_user_id == callback.from_user.id).order_by(ScheduleItem.date)
        if module_id:
            stmt = stmt.where(ScheduleItem.module_id == module_id)
        days = sorted({x.date for x in (await session.scalars(stmt)).all()})
    kb = InlineKeyboardBuilder()
    for d in days:
        kb.button(text=d.strftime("%d.%m.%Y"), callback_data=f"schedule_day:{module_id or 0}:{d.isoformat()}")
    if module_id:
        kb.button(text="➕ Добавить слот вручную", callback_data=f"manual_schedule:{module_id}")
        kb.button(text="⬅️ К модулю", callback_data=f"module:{module_id}")
    kb.button(text="🏠 Главное меню", callback_data="menu")
    kb.adjust(1)
    await send_or_edit(callback, "📅 Выберите день:" if days else "Расписание пока пустое.", kb)


@router.callback_query(F.data.startswith("schedule_day:"))
async def schedule_day(callback: CallbackQuery) -> None:
    _, module_id_raw, day_raw = callback.data.split(":")
    module_id = int(module_id_raw)
    selected_day = date.fromisoformat(day_raw)
    async with SessionLocal() as session:
        stmt = (
            select(ScheduleItem)
            .join(Module)
            .where(Module.owner_user_id == callback.from_user.id, ScheduleItem.date == selected_day)
            .options(selectinload(ScheduleItem.speaker), selectinload(ScheduleItem.module))
            .order_by(ScheduleItem.start_time)
        )
        if module_id:
            stmt = stmt.where(ScheduleItem.module_id == module_id)
        items = (await session.scalars(stmt)).all()
        lines = [f"📅 <b>{selected_day:%d.%m.%Y}</b>", ""]
        for item in items:
            ms = await get_module_speaker(session, item.module_id, item.speaker_id) if item.speaker_id else None
            lines += [
                f"<b>{time_label(item.start_time)}-{time_label(item.end_time)}</b>",
                f"🎤 {html_escape(item.speaker.full_name if item.speaker else item.title)}",
                f"Тема: {html_escape(item.title)}",
                f"Сетап: {html_escape(', '.join(ms.setup) if ms and ms.setup else 'не указан')}",
            ]
            if ms and ms.amount:
                lines.append(f"Сумма: {money(ms.amount)}")
            if ms and (ms.flower_required or ms.speaker.gender.lower().startswith("жен")):
                lines.append("🌸 Нужны цветы")
            lines.append("")
    await send_or_edit(callback, "\n".join(lines), kb_back(module_id or None))


async def recreate_reminders_for_module(session: AsyncSession, module_id: int, user_id: int) -> None:
    settings = await get_settings(session, user_id)
    tz = ZoneInfo(settings.timezone)
    for old in (await session.scalars(select(Reminder).where(Reminder.module_id == module_id, Reminder.is_sent == False))).all():
        await session.delete(old)
    items = (
        await session.scalars(
            select(ScheduleItem)
            .where(ScheduleItem.module_id == module_id, ScheduleItem.speaker_id.is_not(None))
            .options(selectinload(ScheduleItem.speaker))
        )
    ).all()
    for item in items:
        if not item.start_time:
            continue
        start_dt = datetime.combine(item.date, item.start_time).replace(tzinfo=tz).astimezone(ZoneInfo("UTC")).replace(tzinfo=None)
        end_dt = datetime.combine(item.date, item.end_time or item.start_time).replace(tzinfo=tz).astimezone(ZoneInfo("UTC")).replace(tzinfo=None)
        ms = await get_module_speaker(session, item.module_id, item.speaker_id)
        session.add(Reminder(module_id=item.module_id, speaker_id=item.speaker_id, schedule_item_id=item.id, reminder_type="performance", remind_at=start_dt - timedelta(minutes=settings.performance_reminder_minutes), target_user_id=user_id))
        if ms and ms.is_paid:
            session.add(Reminder(module_id=item.module_id, speaker_id=item.speaker_id, schedule_item_id=item.id, reminder_type="documents", remind_at=start_dt - timedelta(minutes=settings.docs_reminder_minutes), target_user_id=user_id))
        if ms and (ms.flower_required or ms.speaker.gender.lower().startswith("жен")) and item.end_time:
            session.add(Reminder(module_id=item.module_id, speaker_id=item.speaker_id, schedule_item_id=item.id, reminder_type="flowers", remind_at=end_dt - timedelta(minutes=settings.flowers_reminder_minutes), target_user_id=user_id))


async def send_due_reminders(bot: Bot) -> None:
    async with SessionLocal() as session:
        now = datetime.utcnow()
        reminders = (
            await session.scalars(
                select(Reminder)
                .where(Reminder.is_sent == False, Reminder.remind_at <= now)
                .order_by(Reminder.remind_at)
                .limit(25)
            )
        ).all()
        for rem in reminders:
            text = await build_reminder_text(session, rem)
            try:
                await bot.send_message(rem.target_user_id, text)
                rem.is_sent = True
                rem.sent_at = datetime.utcnow()
            except Exception as exc:
                log.warning("cannot send reminder %s: %s", rem.id, exc)
        await session.commit()


async def build_reminder_text(session: AsyncSession, rem: Reminder) -> str:
    module = await session.get(Module, rem.module_id)
    speaker = await session.get(Speaker, rem.speaker_id) if rem.speaker_id else None
    item = await session.get(ScheduleItem, rem.schedule_item_id) if rem.schedule_item_id else None
    ms = await get_module_speaker(session, rem.module_id, rem.speaker_id) if rem.speaker_id else None
    if rem.reminder_type == "documents" and ms:
        ds = ms.document_status
        lines = [
            "📄 <b>Проверь документы по платному спикеру</b>",
            "",
            f"👤 {html_escape(speaker.full_name)}",
            f"⏰ Выступление через 30 минут",
            f"💵 Сумма: {money(ms.amount)}",
            f"📄 Тип оформления: {html_escape(ms.contract_type)}",
            "",
            "Статус документов:",
        ]
        for field, title in DOC_FIELDS:
            lines.append(f"{'✅' if ds and getattr(ds, field) else '❌'} {title}")
        lines.append(f"\nПрогресс: {doc_progress(ds)}%")
        return "\n".join(lines)
    if rem.reminder_type == "flowers":
        return "\n".join([
            "🌸 <b>Через 30 минут нужно вынести цветы</b>",
            "",
            f"👤 Спикер: {html_escape(speaker.full_name if speaker else '')}",
            f"📁 Модуль: {html_escape(module.name if module else '')}",
            f"⏰ Выступление заканчивается в {time_label(item.end_time) if item else '--:--'}",
            f"📍 Зал: {html_escape(item.location if item else '')}",
            "",
            "Не забудь подготовить букет и передать его в зал.",
        ])
    return "\n".join([
        "🔔 <b>Через 40 минут выступает спикер</b>",
        "",
        f"👤 {html_escape(speaker.full_name if speaker else '')}",
        f"📁 Модуль: {html_escape(module.name if module else '')}",
        f"⏰ Время: {time_label(item.start_time) if item else '--:--'}-{time_label(item.end_time) if item else '--:--'}",
        f"🎤 Тема: {html_escape(item.title if item else (ms.topic if ms else ''))}",
        f"📍 Зал: {html_escape(item.location if item else '')}",
        "",
        "⚙️ Сетап:",
        *(f"- {html_escape(x)}" for x in (ms.setup if ms and ms.setup else ["не указан"])),
        "",
        f"📞 Контакт: {html_escape(speaker.phone if speaker else '')}",
        f"💬 Telegram: {html_escape(speaker.telegram if speaker else '')}",
        f"💰 Сумма: {money(ms.amount if ms else 0)}",
    ])


async def post_module_document_reminders(bot: Bot) -> None:
    async with SessionLocal() as session:
        users = (await session.scalars(select(AppSetting))).all()
        today = datetime.utcnow().date()
        for settings in users:
            if not settings.post_module_reminders_enabled:
                continue
            modules = (
                await session.scalars(
                    select(Module).where(Module.owner_user_id == settings.user_id, Module.end_date.is_not(None), Module.end_date < today)
                )
            ).all()
            for module in modules:
                if (today - module.end_date).days % max(settings.post_module_interval_days, 1) != 0:
                    continue
                links = (
                    await session.scalars(
                        select(ModuleSpeaker)
                        .where(ModuleSpeaker.module_id == module.id, ModuleSpeaker.is_paid == True)
                        .options(selectinload(ModuleSpeaker.speaker), selectinload(ModuleSpeaker.document_status))
                    )
                ).all()
                for ms in links:
                    if doc_progress(ms.document_status) >= 100:
                        continue
                    missing = [title for field, title in DOC_FIELDS if not ms.document_status or not getattr(ms.document_status, field)]
                    text = "\n".join([
                        "📄 <b>Не закрыты документы по спикеру</b>",
                        "",
                        f"👤 {html_escape(ms.speaker.full_name)}",
                        f"📁 Модуль: {html_escape(module.name)}",
                        f"💵 Сумма: {money(ms.amount)}",
                        f"📄 Тип оформления: {html_escape(ms.contract_type)}",
                        "",
                        "Не закрыто:",
                        *(f"- {html_escape(x)}" for x in missing),
                        "",
                        f"Прогресс: {doc_progress(ms.document_status)}%",
                    ])
                    try:
                        await bot.send_message(settings.user_id, text)
                    except Exception as exc:
                        log.warning("post module reminder failed: %s", exc)
        await session.commit()


@router.callback_query(F.data == "reminders")
async def reminders_view(callback: CallbackQuery) -> None:
    async with SessionLocal() as session:
        reminders = (
            await session.scalars(
                select(Reminder)
                .where(Reminder.target_user_id == callback.from_user.id, Reminder.is_sent == False)
                .order_by(Reminder.remind_at)
                .limit(20)
            )
        ).all()
    lines = ["🔔 <b>Ближайшие напоминания</b>", ""]
    for rem in reminders:
        lines.append(f"{rem.remind_at:%d.%m.%Y %H:%M} UTC | {rem.reminder_type}")
    await send_or_edit(callback, "\n".join(lines) if reminders else "Активных напоминаний пока нет.", kb_main())


@router.callback_query(F.data.startswith("recreate_reminders:"))
async def recreate_reminders_callback(callback: CallbackQuery) -> None:
    ms_id = int(callback.data.split(":")[1])
    async with SessionLocal() as session:
        ms = await session.get(ModuleSpeaker, ms_id)
        if ms:
            await recreate_reminders_for_module(session, ms.module_id, callback.from_user.id)
            await session.commit()
    await callback.answer("Напоминания обновлены", show_alert=True)


@router.callback_query(F.data == "settings")
async def settings_view(callback: CallbackQuery) -> None:
    async with SessionLocal() as session:
        s = await get_settings(session, callback.from_user.id)
        await session.commit()
    text = "\n".join([
        "⚙️ <b>Настройки</b>",
        "",
        f"Часовой пояс: {html_escape(s.timezone)}",
        f"Напоминание перед выступлением: {s.performance_reminder_minutes} мин.",
        f"Напоминание о документах: {s.docs_reminder_minutes} мин.",
        f"Напоминание о цветах: {s.flowers_reminder_minutes} мин. до конца",
        f"ИИ-помощник: {'включен' if s.ai_enabled else 'выключен'}",
        f"Напоминания после модуля: {'включены' if s.post_module_reminders_enabled else 'выключены'}",
        f"Интервал после модуля: каждые {s.post_module_interval_days} дн.",
    ])
    kb = InlineKeyboardBuilder()
    buttons = [
        ("Часовой пояс", "settings_edit:timezone"),
        ("Минуты до выступления", "settings_edit:performance_minutes"),
        ("Документы: 20/30", "settings_docs_toggle"),
        ("Минуты до цветов", "settings_edit:flowers_minutes"),
        ("ИИ вкл/выкл", "settings_toggle_ai"),
        ("После модуля вкл/выкл", "settings_toggle_post"),
        ("Интервал после модуля", "settings_edit:post_interval"),
        ("🏠 Главное меню", "menu"),
    ]
    for t, d in buttons:
        kb.button(text=t, callback_data=d)
    kb.adjust(2, 2, 2, 1, 1)
    await send_or_edit(callback, text, kb)


@router.callback_query(F.data == "settings_docs_toggle")
async def settings_docs_toggle(callback: CallbackQuery) -> None:
    async with SessionLocal() as session:
        s = await get_settings(session, callback.from_user.id)
        s.docs_reminder_minutes = 20 if s.docs_reminder_minutes == 30 else 30
        await session.commit()
    await settings_view(callback)


@router.callback_query(F.data.in_({"settings_toggle_ai", "settings_toggle_post"}))
async def settings_toggle(callback: CallbackQuery) -> None:
    async with SessionLocal() as session:
        s = await get_settings(session, callback.from_user.id)
        if callback.data == "settings_toggle_ai":
            s.ai_enabled = not s.ai_enabled
        else:
            s.post_module_reminders_enabled = not s.post_module_reminders_enabled
        await session.commit()
    await settings_view(callback)


@router.callback_query(F.data.startswith("settings_edit:"))
async def settings_edit(callback: CallbackQuery, state: FSMContext) -> None:
    field = callback.data.split(":")[1]
    await state.update_data(field=field)
    await state.set_state(getattr(SettingsEdit, field))
    await callback.message.edit_text("Введите новое значение:")
    await callback.answer()


@router.message(SettingsEdit.timezone)
@router.message(SettingsEdit.performance_minutes)
@router.message(SettingsEdit.flowers_minutes)
@router.message(SettingsEdit.post_interval)
async def settings_save(message: Message, state: FSMContext) -> None:
    data = await state.get_data()
    field = data["field"]
    value = message.text.strip()
    async with SessionLocal() as session:
        s = await get_settings(session, message.from_user.id)
        if field == "timezone":
            ZoneInfo(value)
            s.timezone = value
        else:
            setattr(s, field, max(1, int(re.sub(r"\D", "", value) or 1)))
        await session.commit()
    await state.clear()
    await message.answer("✅ Настройки сохранены.", reply_markup=kb_main().as_markup())


async def parse_text_command(message_text: str) -> dict[str, Any] | None:
    if not openai_client or not OPENAI_API_KEY:
        return heuristic_text_command(message_text)
    schema = {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "intent": {"type": "string", "enum": ["create_speaker", "update_speaker_payment", "update_document_status", "update_setup", "update_schedule_time", "unknown"]},
            "speaker_name": {"type": "string"},
            "module_name": {"type": "string"},
            "fields": {"type": "object", "additionalProperties": True},
            "needs_confirmation": {"type": "boolean"},
            "summary": {"type": "string"},
        },
        "required": ["intent", "speaker_name", "module_name", "fields", "needs_confirmation", "summary"],
    }
    try:
        response = await openai_client.responses.create(
            model=OPENAI_MODEL,
            input=[
                {"role": "system", "content": "Преобразуй сообщение менеджера мероприятий в структурированную команду для бота. Важные изменения помечай needs_confirmation=true."},
                {"role": "user", "content": message_text},
            ],
            text={"format": {"type": "json_schema", "name": "bot_command", "schema": schema, "strict": True}},
        )
        return json.loads(response.output_text)
    except Exception as exc:
        log.exception("AI command parse failed: %s", exc)
        return heuristic_text_command(message_text)


def heuristic_text_command(text: str) -> dict[str, Any] | None:
    low = text.lower()
    amount = re.search(r"(\d[\d\s]{3,})", text)
    name = ""
    m = re.search(r"(иванов[а-яё\s]*|петров[а-яё\s]*|смирнов[а-яё\s]*|[А-ЯЁ][а-яё]+(?:\s+[А-ЯЁ][а-яё]+){0,2})", text)
    if m:
        name = m.group(1).strip()
    if "сумм" in low or amount:
        return {"intent": "update_speaker_payment", "speaker_name": name, "module_name": "", "fields": {"amount": int(re.sub(r"\D", "", amount.group(1))) if amount else 0}, "needs_confirmation": True, "summary": "изменить сумму"}
    if "подписан" in low or "счет" in low or "чек" in low or "договор" in low:
        fields = {}
        if "подпис" in low:
            fields["contract_act_signed"] = True
        if "счет" in low or "чек" in low:
            fields["invoice_or_receipt_received"] = True
        if "соглас" in low:
            fields["contract_act_approved"] = True
        return {"intent": "update_document_status", "speaker_name": name, "module_name": "", "fields": fields, "needs_confirmation": True, "summary": "обновить документы"}
    if "цвет" in low or "кликер" in low or "презентац" in low or "микрофон" in low:
        return {"intent": "update_setup", "speaker_name": name, "module_name": "", "fields": {"setup_add": split_setup(text)}, "needs_confirmation": True, "summary": "обновить сетап"}
    if "добав" in low and name:
        return {"intent": "create_speaker", "speaker_name": name, "module_name": "", "fields": {}, "needs_confirmation": True, "summary": "добавить спикера"}
    return None


async def apply_text_command(session: AsyncSession, user_id: int, command: dict[str, Any]) -> str:
    intent = command.get("intent")
    speaker = await find_speaker(session, user_id, command.get("speaker_name", ""))
    module = await find_module(session, user_id, command.get("module_name", ""))
    fields = command.get("fields") or {}
    if intent == "create_speaker":
        if not module:
            return "Не нашел модуль. Создайте модуль или укажите его название."
        speaker = Speaker(full_name=command.get("speaker_name") or fields.get("full_name") or "Новый спикер", gender=fields.get("gender", ""), phone=fields.get("phone", ""), telegram=fields.get("telegram", ""), email=fields.get("email", ""), owner_user_id=user_id)
        session.add(speaker)
        await session.flush()
        ms = await ensure_module_speaker(session, module, speaker, topic=fields.get("topic", ""), setup=split_setup(fields.get("setup", [])), is_paid=bool(fields.get("amount")), amount=float(fields.get("amount") or 0), contract_type=fields.get("contract_type", ""))
        await log_action(session, user_id, "ai_create_speaker", "speaker", speaker.id, new=fields)
        await session.commit()
        return f"✅ Добавил спикера: {speaker.full_name}. Карточка доступна в модуле «{module.name}»."
    if not speaker:
        return "Не нашел спикера. Уточните ФИО."
    if not module:
        link = await session.scalar(
            select(ModuleSpeaker)
            .where(ModuleSpeaker.speaker_id == speaker.id)
            .options(selectinload(ModuleSpeaker.module))
        )
        module = link.module if link else None
    if not module:
        return "Не нашел модуль спикера."
    ms = await get_module_speaker(session, module.id, speaker.id)
    if intent == "update_speaker_payment":
        old = {"amount": ms.amount}
        ms.amount = float(fields.get("amount") or 0)
        ms.is_paid = ms.amount > 0
        if fields.get("contract_type"):
            ms.contract_type = fields["contract_type"]
        if ms.is_paid and not ms.document_status:
            session.add(DocumentStatus(module_speaker_id=ms.id))
        await log_action(session, user_id, "ai_update_payment", "module_speaker", ms.id, old=old, new=fields)
        await session.commit()
        return f"✅ Обновил сумму для {speaker.full_name}: {money(ms.amount)}."
    if intent == "update_document_status":
        if not ms.document_status:
            ms.document_status = DocumentStatus(module_speaker_id=ms.id)
        for field in fields:
            if field in dict(DOC_FIELDS):
                setattr(ms.document_status, field, bool(fields[field]))
        doc_progress(ms.document_status)
        await log_action(session, user_id, "ai_update_docs", "document_status", ms.document_status.id, new=fields)
        await session.commit()
        return f"✅ Обновил документы для {speaker.full_name}. Прогресс: {ms.document_status.progress_percent}%."
    if intent == "update_setup":
        additions = split_setup(fields.get("setup_add") or fields.get("setup") or [])
        ms.setup = sorted(set((ms.setup or []) + additions))
        ms.flower_required = ms.flower_required or "цветы" in ms.setup or bool(fields.get("flower_required"))
        await log_action(session, user_id, "ai_update_setup", "module_speaker", ms.id, new=fields)
        await session.commit()
        return f"✅ Обновил сетап для {speaker.full_name}: {', '.join(ms.setup)}."
    if intent == "update_schedule_time":
        item = await session.scalar(select(ScheduleItem).where(ScheduleItem.speaker_id == speaker.id, ScheduleItem.module_id == module.id))
        if not item:
            return "Не нашел слот расписания для этого спикера."
        item.start_time = parse_time(str(fields.get("start_time", ""))) or item.start_time
        item.end_time = parse_time(str(fields.get("end_time", ""))) or item.end_time
        await recreate_reminders_for_module(session, module.id, user_id)
        await log_action(session, user_id, "ai_update_schedule", "schedule_item", item.id, new=fields)
        await session.commit()
        return f"✅ Перенес выступление {speaker.full_name} на {time_label(item.start_time)}."
    return "Не понял действие. Попробуйте написать чуть конкретнее."


@router.callback_query(F.data.in_({"ai_confirm", "ai_cancel"}))
async def ai_confirm(callback: CallbackQuery) -> None:
    command = PENDING_AI_ACTIONS.pop(callback.from_user.id, None)
    if callback.data == "ai_cancel" or not command:
        await callback.message.edit_text("Отменено.", reply_markup=kb_main().as_markup())
        await callback.answer()
        return
    async with SessionLocal() as session:
        result = await apply_text_command(session, callback.from_user.id, command)
    await callback.message.edit_text(result, reply_markup=kb_main().as_markup())
    await callback.answer()


@router.message(F.text)
async def natural_text(message: Message, state: FSMContext) -> None:
    current = await state.get_state()
    if current:
        return
    async with SessionLocal() as session:
        settings = await get_settings(session, message.from_user.id)
        if not settings.ai_enabled:
            await message.answer("ИИ-помощник выключен в настройках.", reply_markup=kb_main().as_markup())
            return
    command = await parse_text_command(message.text)
    if not command or command.get("intent") == "unknown":
        await message.answer("Не понял задачу. Можно нажать кнопку в меню или написать подробнее.", reply_markup=kb_main().as_markup())
        return
    if command.get("needs_confirmation", True):
        PENDING_AI_ACTIONS[message.from_user.id] = command
        kb = InlineKeyboardBuilder()
        kb.button(text="✅ Да, выполнить", callback_data="ai_confirm")
        kb.button(text="❌ Отмена", callback_data="ai_cancel")
        kb.adjust(1)
        await message.answer(f"Я понял задачу:\n\n{html_escape(command.get('summary') or json.dumps(command, ensure_ascii=False))}\n\nПодтвердить изменение?", reply_markup=kb.as_markup())
    else:
        async with SessionLocal() as session:
            result = await apply_text_command(session, message.from_user.id, command)
        await message.answer(result, reply_markup=kb_main().as_markup())


async def init_db() -> None:
    async with engine.begin() as conn:
        await conn.run_sync(Base.metadata.create_all)


async def main() -> None:
    if not BOT_TOKEN:
        raise RuntimeError("BOT_TOKEN не задан. Создайте .env и укажите BOT_TOKEN.")
    await init_db()
    bot = Bot(BOT_TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
    dp = Dispatcher(storage=MemoryStorage())
    dp.include_router(router)
    scheduler = AsyncIOScheduler(timezone="UTC")
    scheduler.add_job(send_due_reminders, "interval", seconds=60, args=[bot], id="send_due_reminders")
    scheduler.add_job(post_module_document_reminders, "interval", hours=6, args=[bot], id="post_module_document_reminders")
    scheduler.start()
    log.info("Bot started")
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
