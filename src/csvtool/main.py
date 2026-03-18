"""CLI entry point for CSV Automation Tool."""

import asyncio
import sys
from pathlib import Path
from typing import Optional

import click

from csvtool import __version__
from csvtool.evidence import EvidenceGenerator
from csvtool.executor import Executor, ExecutorConfig


@click.group()
@click.version_option(version=__version__, prog_name="CSV Automation Tool")
def cli():
    """CSV Automation Tool - Automated test execution with AI-powered navigation."""
    pass


@cli.command()
@click.option(
    "--url",
    required=True,
    help="Application URL to test"
)
@click.option(
    "--workbook",
    required=True,
    type=click.Path(exists=True, path_type=Path),
    help="Path to test scenarios Excel workbook"
)
@click.option(
    "--output",
    type=click.Path(path_type=Path),
    default=Path("output"),
    help="Output directory for evidence documents"
)
@click.option(
    "--cache-dir",
    type=click.Path(path_type=Path),
    default=Path("cache"),
    help="Directory for navigation cache"
)
@click.option(
    "--headless/--no-headless",
    default=True,
    help="Run browser in headless mode"
)
@click.option(
    "--stop-on-failure/--continue-on-failure",
    default=True,
    help="Stop execution on first failure"
)
@click.option(
    "--api-key",
    envvar="ANTHROPIC_API_KEY",
    help="Anthropic API key (or set ANTHROPIC_API_KEY env var)"
)
@click.option(
    "--hybrid-llm/--claude-only",
    default=True,
    help="Use hybrid mode (local LLM + Claude fallback) or Claude only"
)
@click.option(
    "--ollama-url",
    default="http://localhost:11434",
    help="Ollama server URL for local LLM"
)
@click.option(
    "--ollama-model",
    default="llava:13b",
    help="Ollama model to use (e.g., llava:13b, minicpm-v:8b)"
)
def run(
    url: str,
    workbook: Path,
    output: Path,
    cache_dir: Path,
    headless: bool,
    stop_on_failure: bool,
    api_key: Optional[str],
    hybrid_llm: bool,
    ollama_url: str,
    ollama_model: str
):
    """Run CSV test scenarios against an application."""
    click.echo(f"CSV Automation Tool v{__version__}")
    click.echo(f"URL: {url}")
    click.echo(f"Workbook: {workbook}")
    click.echo()

    # Create config
    config = ExecutorConfig(
        url=url,
        workbook_path=workbook,
        output_dir=output,
        cache_dir=cache_dir,
        headless=headless,
        stop_on_failure=stop_on_failure,
        api_key=api_key,
        use_hybrid_llm=hybrid_llm,
        ollama_url=ollama_url,
        ollama_model=ollama_model
    )

    # Run executor
    click.echo("Starting test execution...")
    executor = Executor(config)

    try:
        summary = asyncio.run(executor.run())
    except Exception as e:
        click.echo(f"Execution error: {e}", err=True)
        sys.exit(1)

    # Generate evidence
    click.echo("Generating evidence document...")
    try:
        generator = EvidenceGenerator(output)
        doc_path = generator.generate(summary)
        click.echo(f"✅ Evidence document created: {doc_path}")
        click.echo(f"   File size: {doc_path.stat().st_size} bytes")
    except Exception as e:
        click.echo(f"❌ Failed to generate evidence document: {e}", err=True)
        click.echo(f"   Output directory: {output.absolute()}", err=True)
        sys.exit(1)

    # Print summary
    click.echo()
    click.echo("=" * 50)
    click.echo("EXECUTION SUMMARY")
    click.echo("=" * 50)
    click.echo(f"Total Tests: {summary.total_tests}")
    click.echo(f"Passed: {summary.passed_tests}")
    click.echo(f"Failed: {summary.failed_tests}")
    click.echo(f"Success Rate: {summary.success_rate:.1f}%")
    click.echo()
    click.echo(f"Evidence Document: {doc_path}")

    # Exit with appropriate code
    if summary.failed_tests > 0:
        sys.exit(1)
    sys.exit(0)


if __name__ == "__main__":
    cli()
