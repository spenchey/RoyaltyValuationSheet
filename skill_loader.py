#!/usr/bin/env python3
"""
Skill Loader Framework for Music Royalty Valuation Tool
Implements a modular skill-based architecture inspired by Claude Skills.
Skills are self-contained modules that extend valuation capabilities.
"""

import os
import yaml
import importlib.util
from typing import Dict, List, Optional, Any, Callable
from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class SkillMetadata:
    """Metadata for a skill extracted from SKILL.md frontmatter."""
    name: str
    description: str
    version: str = "1.0.0"
    author: str = ""
    dependencies: List[str] = field(default_factory=list)
    triggers: List[str] = field(default_factory=list)


@dataclass
class Skill:
    """Represents a loaded skill with its metadata and executor."""
    metadata: SkillMetadata
    path: Path
    executor: Optional[Callable] = None
    instructions: str = ""
    is_loaded: bool = False


class SkillRegistry:
    """Registry for managing and discovering skills."""

    def __init__(self, skills_dir: Optional[str] = None):
        if skills_dir is None:
            # Default to skills directory relative to this file
            self.skills_dir = Path(__file__).parent / "skills"
        else:
            self.skills_dir = Path(skills_dir)

        self.skills: Dict[str, Skill] = {}
        self._discover_skills()

    def _discover_skills(self) -> None:
        """Discover all available skills in the skills directory."""
        if not self.skills_dir.exists():
            return

        for skill_path in self.skills_dir.iterdir():
            if skill_path.is_dir() and not skill_path.name.startswith('.'):
                skill_md = skill_path / "SKILL.md"
                if skill_md.exists():
                    try:
                        metadata = self._parse_skill_metadata(skill_md)
                        self.skills[metadata.name] = Skill(
                            metadata=metadata,
                            path=skill_path,
                            is_loaded=False
                        )
                    except Exception as e:
                        print(f"Warning: Could not load skill from {skill_path}: {e}")

    def _parse_skill_metadata(self, skill_md_path: Path) -> SkillMetadata:
        """Parse SKILL.md frontmatter to extract metadata."""
        content = skill_md_path.read_text()

        # Extract YAML frontmatter
        if content.startswith('---'):
            parts = content.split('---', 2)
            if len(parts) >= 3:
                frontmatter = yaml.safe_load(parts[1])
                return SkillMetadata(
                    name=frontmatter.get('name', skill_md_path.parent.name),
                    description=frontmatter.get('description', ''),
                    version=frontmatter.get('version', '1.0.0'),
                    author=frontmatter.get('author', ''),
                    dependencies=frontmatter.get('dependencies', []),
                    triggers=frontmatter.get('triggers', [])
                )

        # Fallback: use directory name
        return SkillMetadata(
            name=skill_md_path.parent.name,
            description=""
        )

    def load_skill(self, skill_name: str) -> Optional[Skill]:
        """Load a skill's full instructions and executor."""
        if skill_name not in self.skills:
            return None

        skill = self.skills[skill_name]
        if skill.is_loaded:
            return skill

        # Load instructions from SKILL.md
        skill_md = skill.path / "SKILL.md"
        if skill_md.exists():
            content = skill_md.read_text()
            # Extract body (after frontmatter)
            if content.startswith('---'):
                parts = content.split('---', 2)
                if len(parts) >= 3:
                    skill.instructions = parts[2].strip()

        # Load executor from executor.py if exists
        executor_path = skill.path / "executor.py"
        if executor_path.exists():
            spec = importlib.util.spec_from_file_location(
                f"skills.{skill_name}.executor",
                executor_path
            )
            if spec and spec.loader:
                module = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(module)
                if hasattr(module, 'execute'):
                    skill.executor = module.execute

        skill.is_loaded = True
        return skill

    def get_skill(self, skill_name: str) -> Optional[Skill]:
        """Get a skill by name, loading it if necessary."""
        return self.load_skill(skill_name)

    def list_skills(self) -> List[SkillMetadata]:
        """List all available skills (metadata only)."""
        return [skill.metadata for skill in self.skills.values()]

    def find_relevant_skills(self, query: str) -> List[Skill]:
        """Find skills relevant to a query based on triggers and description."""
        query_lower = query.lower()
        relevant = []

        for skill in self.skills.values():
            # Check triggers
            for trigger in skill.metadata.triggers:
                if trigger.lower() in query_lower:
                    relevant.append(skill)
                    break
            else:
                # Check description
                if any(word in skill.metadata.description.lower()
                       for word in query_lower.split()):
                    relevant.append(skill)

        return relevant

    def execute_skill(self, skill_name: str, **kwargs) -> Any:
        """Execute a skill with given parameters."""
        skill = self.load_skill(skill_name)
        if skill is None:
            raise ValueError(f"Skill '{skill_name}' not found")

        if skill.executor is None:
            raise ValueError(f"Skill '{skill_name}' has no executor")

        return skill.executor(**kwargs)


class SkillPipeline:
    """Pipeline for executing multiple skills in sequence."""

    def __init__(self, registry: SkillRegistry):
        self.registry = registry
        self.steps: List[tuple] = []  # (skill_name, params_mapper)

    def add_step(self, skill_name: str, params_mapper: Optional[Callable] = None):
        """Add a skill step to the pipeline."""
        self.steps.append((skill_name, params_mapper))
        return self

    def execute(self, initial_data: Dict[str, Any]) -> Dict[str, Any]:
        """Execute all pipeline steps."""
        data = initial_data.copy()

        for skill_name, params_mapper in self.steps:
            # Map parameters if mapper provided
            if params_mapper:
                params = params_mapper(data)
            else:
                params = data

            # Execute skill
            result = self.registry.execute_skill(skill_name, **params)

            # Merge result into data
            if isinstance(result, dict):
                data.update(result)
            else:
                data[skill_name + '_result'] = result

        return data


# Convenience function to create a global registry
_global_registry: Optional[SkillRegistry] = None

def get_registry() -> SkillRegistry:
    """Get or create the global skill registry."""
    global _global_registry
    if _global_registry is None:
        _global_registry = SkillRegistry()
    return _global_registry


def register_skill(skill_path: str) -> None:
    """Register a skill from an external path."""
    registry = get_registry()
    skill_md = Path(skill_path) / "SKILL.md"
    if skill_md.exists():
        metadata = registry._parse_skill_metadata(skill_md)
        registry.skills[metadata.name] = Skill(
            metadata=metadata,
            path=Path(skill_path),
            is_loaded=False
        )
