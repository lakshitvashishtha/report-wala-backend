from .base_engine import BaseReportEngine
from .academic_project_report_engine import AcademicProjectEngine
from .phd_thesis_engine import PhdThesisEngine
from .doctoral_dissertation_engine import DoctoralDissertationEngine
from .masters_dissertation_engine import MastersDissertationEngine
from .research_paper_engine import ResearchPaperEngine
from .literature_review_engine import LiteratureReviewEngine
from .literature_report_engine import LiteratureReportEngine
from .btech_project_report_engine import BtechProjectEngine
from .internship_report_engine import InternshipEngine
from .psd_report_engine import PsdReportEngine
from .vsd_report_engine import VsdReportEngine
from .technical_report_engine import TechnicalReportEngine
from .business_report_engine import BusinessReportEngine
from .medical_case_study_engine import MedicalCaseStudyEngine
from .professional_project_report_engine import ProfessionalProjectEngine
from .training_report_engine import TrainingEngine
from .seminar_report_engine import SeminarEngine
from .feasibility_report_engine import FeasibilityEngine
from .policy_report_engine import PolicyEngine

def get_engine_for_report(report_type, metadata, job_dir=None, status_file=None):
    rt_lower = str(report_type).lower().strip()

    if rt_lower == "academic project report":
        return AcademicProjectEngine(metadata, job_dir, status_file)
    if rt_lower == "phd thesis":
        return PhdThesisEngine(metadata, job_dir, status_file)
    if rt_lower == "doctoral dissertation":
        return DoctoralDissertationEngine(metadata, job_dir, status_file)
    if rt_lower == "master's dissertation":
        return MastersDissertationEngine(metadata, job_dir, status_file)
    if rt_lower == "research paper":
        return ResearchPaperEngine(metadata, job_dir, status_file)
    if rt_lower == "literature review":
        return LiteratureReviewEngine(metadata, job_dir, status_file)
    if rt_lower == "literature report":
        return LiteratureReportEngine(metadata, job_dir, status_file)
    if rt_lower == "btech project report":
        return BtechProjectEngine(metadata, job_dir, status_file)
    if rt_lower == "internship report":
        return InternshipEngine(metadata, job_dir, status_file)
    if rt_lower == "practice school (psd) report":
        return PsdReportEngine(metadata, job_dir, status_file)
    if rt_lower == "vlsi system design (vsd) report":
        return VsdReportEngine(metadata, job_dir, status_file)
    if rt_lower == "technical report":
        return TechnicalReportEngine(metadata, job_dir, status_file)
    if rt_lower == "business report":
        return BusinessReportEngine(metadata, job_dir, status_file)
    if rt_lower == "medical case study":
        return MedicalCaseStudyEngine(metadata, job_dir, status_file)
    if rt_lower == "professional project report":
        return ProfessionalProjectEngine(metadata, job_dir, status_file)
    if rt_lower == "training report":
        return TrainingEngine(metadata, job_dir, status_file)
    if rt_lower == "seminar report":
        return SeminarEngine(metadata, job_dir, status_file)
    if rt_lower == "feasibility report":
        return FeasibilityEngine(metadata, job_dir, status_file)
    if rt_lower == "policy report":
        return PolicyEngine(metadata, job_dir, status_file)

    # Fallback to the base engine
    return BaseReportEngine(metadata, job_dir, status_file)
