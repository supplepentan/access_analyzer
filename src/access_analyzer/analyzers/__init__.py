"""Cross-object dependency and Python migration analyzers."""

from access_analyzer.analyzers.dependency_analyzer import build_dependency_graph
from access_analyzer.analyzers.migration_analyzer import build_migration_inventory

__all__ = ["build_dependency_graph", "build_migration_inventory"]
