"""Root compatibility wrapper for the workflow CLI module."""

from workflow.vi_report_cli import *  # noqa: F401,F403


if __name__ == "__main__":
    from workflow.vi_report_cli import main

    main()
