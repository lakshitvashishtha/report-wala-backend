from .base_engine import BaseReportEngine

class TrainingEngine(BaseReportEngine):
    def __init__(self, metadata, job_dir, status_file=None):
        super().__init__(metadata, job_dir, status_file)
        self.preferred_models = ["gemini-1.5-pro", "gemini-pro-latest", "gemini-flash-latest"]

    def _report_type_profile(self):
        profile = super()._report_type_profile()
        profile["missing"] += "\nFocus: Learning objectives achieved, skill acquisition tracking, and competency evaluations."
        return profile
