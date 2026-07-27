from .base_engine import BaseReportEngine

class MedicalCaseStudyEngine(BaseReportEngine):
    def __init__(self, metadata, job_dir, status_file=None):
        super().__init__(metadata, job_dir, status_file)
        self.preferred_models = ["gemini-1.5-pro", "gemini-pro-latest", "gemini-flash-latest"]

    def _report_type_profile(self):
        profile = super()._report_type_profile()
        profile["missing"] += "\nFocus: Clinical terminology, patient history, diagnostic processes, and treatment outcomes following medical standards."
        return profile

    def _get_section_label(self):
        return ""

    def _get_citation_style(self):
        return "AMA format (American Medical Association)"

    def _get_in_text_citation_style(self):
        return "superscript numbers like [1] or 1"
