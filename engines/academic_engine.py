from .base_engine import BaseReportEngine

class AcademicEngine(BaseReportEngine):
    def __init__(self, metadata, job_dir, status_file=None):
        super().__init__(metadata, job_dir, status_file)
        # Prioritize pro models for complex academic reports
        self.preferred_models = ["gemini-pro-latest", "gemini-1.5-pro", "gemini-flash-latest", "gemini-flash-lite-latest"]

    def _report_type_profile(self):
        # Override to provide the best academic profile
        profile = super()._report_type_profile()
        profile["missing"] += "\nFocus heavily on rigorous methodology, citations, and critical analysis suitable for a top-tier academic institution."
        return profile
