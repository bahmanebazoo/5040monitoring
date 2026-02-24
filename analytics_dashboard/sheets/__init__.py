"""
Auto-discovery: همه شیت‌ساز‌ها رو جمع می‌کنه.
OCP: شیت جدید؟ فقط فایل بساز. اینجا خودکار پیدا میشه.
"""

from ..style_manager import StyleManager
from .base import SheetCreator

from .kpi_sheet import KPISheetCreator
from .hourly_sheet import HourlySheetCreator
from .agent_sheet import AgentSheetCreator
from .matching_sheet import MatchingSheetCreator
from .confidence_sheet import ConfidenceSheetCreator
from .wait_time_sheet import WaitTimeSheetCreator
from .status_sheet import StatusSheetCreator
from .delta_sheet import DeltaSheetCreator
from .sla_sheet import SLASheetCreator
from .raw_data_sheet import RawDataSheetCreator
from .work_hours_sheet import WorkHoursSheetCreator

# ترتیب پیش‌فرض شیت‌ها
DEFAULT_CREATORS: list[type[SheetCreator]] = [
    KPISheetCreator,
    HourlySheetCreator,
    AgentSheetCreator,
    WorkHoursSheetCreator,
    ConfidenceSheetCreator,
    MatchingSheetCreator,
    WaitTimeSheetCreator,
    StatusSheetCreator,
    DeltaSheetCreator,
    SLASheetCreator,
    RawDataSheetCreator,
]


def build_default_creators(style: StyleManager) -> list[SheetCreator]:
    """instance بساز از تمام شیت‌سازها."""
    return [cls(style) for cls in DEFAULT_CREATORS]
