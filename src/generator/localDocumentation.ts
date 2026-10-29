import type { ParsedSolution } from '../types/solution';
import { collectMissingDependencies, missingDependenciesNotice, MISSING_DEPENDENCIES_NOTE } from './missingDependencies';
import { heading, mdEscape } from './markdownFormatting';

export function generateDiagnosticsMarkdown(solution: ParsedSolution): string {
  const issues = [...new Set([
    ...solution.warnings,
    ...(solution.missingDependenciesError ? [solution.missingDependenciesError] : []),
    ...solution.environmentVariables.flatMap((variable) => variable.currentValueError ? [`${variable.schemaName}: ${variable.currentValueError}`] : []),
  ])];
  if (issues.length === 0) return '';
  return [heading(1, `${solution.metadata.displayName || solution.metadata.uniqueName}: Diagnostics`), '',
    heading(2, 'Parse Warnings and Issues'), '', ...issues.map((issue) => `- ${mdEscape(issue)}`), ''].join('\n');
}

export function generateDependencyReportMarkdown(solution: ParsedSolution): string {
  const dependencies = collectMissingDependencies(solution);
  const notice = missingDependenciesNotice(solution);
  return [heading(1, `${solution.metadata.displayName || solution.metadata.uniqueName}: Missing Dependencies`), '',
    MISSING_DEPENDENCIES_NOTE, '',
    ...(notice ? [notice] : dependencies.length === 0 ? ['No missing import dependencies declared in the solution ZIP.'] : [
      '| Solution | Dependent Item | Item Type | Required Component | Required Type | Tables |',
      '| --- | --- | --- | --- | --- | --- |',
      ...dependencies.map((dependency) => `| ${[
        dependency.solutionUniqueName || solution.metadata.uniqueName, dependency.dependentName, dependency.dependentType,
        [dependency.requiredParent, dependency.requiredName].filter(Boolean).join('.'), dependency.requiredType,
        dependency.tableLogicalNames.join(', '),
      ].map(mdEscape).join(' | ')} |`),
    ]), ''].join('\n');
}