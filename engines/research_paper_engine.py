from .base_engine import BaseReportEngine

class ResearchPaperEngine(BaseReportEngine):
    def __init__(self, metadata, job_dir, status_file=None):
        super().__init__(metadata, job_dir, status_file)
        self.preferred_models = ["gemini-2.0-flash", "gemini-1.5-pro", "gemini-pro-latest", "gemini-flash-latest"]

    def _report_type_profile(self):
        profile = super()._report_type_profile()
        profile["missing"] += "\nFocus: Concise, peer-review style formatting, strong data presentation, and clear methodology."
        return profile

    def _get_section_label(self):
        return ""
