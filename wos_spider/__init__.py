from .parser import parse_paper_html

try:
    from .spider import SpiderEngine
except ImportError:
    pass

try:
    from .merger import merge_files
except ImportError:
    pass

try:
    from .visualizer import generate_matplotlib_chart
except ImportError:
    pass
