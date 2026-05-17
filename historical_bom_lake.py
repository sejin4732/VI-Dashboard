"""Root compatibility wrapper for the historical BOM lake CLI."""

from workflow.historical_bom_lake import *  # noqa: F401,F403


if __name__ == "__main__":
    from workflow.historical_bom_lake import main

    main()
