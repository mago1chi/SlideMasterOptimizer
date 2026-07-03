"""Core OOXML analysis and optimization helpers."""

from .pptx_optimizer import (
    AnalysisResult,
    LayoutCandidate,
    MasterCandidate,
    OptimizeResult,
    analyze_pptx,
    optimize_pptx,
)

__all__ = [
    "AnalysisResult",
    "LayoutCandidate",
    "MasterCandidate",
    "OptimizeResult",
    "analyze_pptx",
    "optimize_pptx",
]
