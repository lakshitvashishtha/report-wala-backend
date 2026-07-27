from .base_engine import BaseReportEngine

class TechnicalReportEngine(BaseReportEngine):
    def __init__(self, metadata, job_dir, status_file=None):
        super().__init__(metadata, job_dir, status_file)
        self.preferred_models = ["gemini-1.5-pro", "gemini-pro-latest", "gemini-flash-latest"]

    def _report_type_profile(self):
        profile = super()._report_type_profile()
        profile["missing"] += "\nFocus heavily on technical details, architecture, data structures, algorithms, and objective performance metrics. Maintain a clinical, precise technical tone."
        return profile

    def _get_citation_style(self):
        return "IEEE format"

    def _get_in_text_citation_style(self):
        return "[1], [2]"
