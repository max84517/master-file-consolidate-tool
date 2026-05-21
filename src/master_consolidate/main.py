"""Entry point."""
from master_consolidate.config import settings
from master_consolidate.ui import app


def main():
    cfg = settings.load()
    app.run(cfg, settings.save)


if __name__ == "__main__":
    main()
