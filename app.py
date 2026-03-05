from .ui import (
    BuildMixin, SettingsMixin, AuthUIMixin, EmailLoadingMixin,
    ListRenderMixin, KeyboardMixin, DetailViewMixin, ActionsMixin,
    AttachmentsMixin, ComposeMixin, TrainRulesMixin, MeetingsMixin,
    UtilsMixin,
)

class EmailDashboard(
    BuildMixin, SettingsMixin, AuthUIMixin, EmailLoadingMixin,
    ListRenderMixin, KeyboardMixin, DetailViewMixin, ActionsMixin,
    AttachmentsMixin, ComposeMixin, TrainRulesMixin, MeetingsMixin,
    UtilsMixin,
):
    """Outlook Email Intelligence Dashboard."""
    pass
