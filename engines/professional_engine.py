from .base_engine import BaseReportEngine

class ProfessionalEngine(BaseReportEngine):
    def __init__(self, metadata, job_dir, status_file=None):
        super().__init__(metadata, job_dir, status_file)
        self.preferred_models = ["gemini-pro-latest", "gemini-flash-latest"]

    def _report_type_profile(self):
        profile = super()._report_type_profile()
        profile["missing"] += "\nEnsure the tone is highly professional, concise, and executive-ready. Focus on actionable insights, ROI, and strategic implications."
        return profile
