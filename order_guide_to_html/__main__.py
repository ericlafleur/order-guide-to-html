import os
import sys

if __package__ is None or __package__ == '':
    # python3 order_guide_to_html  — no package context; add parent to path
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    from order_guide_to_html.cli import main
else:
    # python3 -m order_guide_to_html  — package context available
    from .cli import main

raise SystemExit(main())
