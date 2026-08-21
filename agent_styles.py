"""
ANSI colors, spinners, and formatted output for CLI agent.
Handles TTY detection and graceful degradation.
"""
import sys
import time
import os


class Colors:
    """ANSI color codes for terminal output."""
    RESET = "\033[0m"
    RED = "\033[91m"
    GREEN = "\033[92m"
    YELLOW = "\033[93m"
    BLUE = "\033[94m"
    MAGENTA = "\033[95m"
    CYAN = "\033[96m"
    WHITE = "\033[97m"
    BOLD = "\033[1m"
    DIM = "\033[2m"

    @staticmethod
    def disable():
        """Disable all colors."""
        Colors.RESET = ""
        Colors.RED = ""
        Colors.GREEN = ""
        Colors.YELLOW = ""
        Colors.BLUE = ""
        Colors.MAGENTA = ""
        Colors.CYAN = ""
        Colors.WHITE = ""
        Colors.BOLD = ""
        Colors.DIM = ""


# Auto-disable colors if not TTY or NO_COLOR env var set
if not sys.stdout.isatty() or os.environ.get('NO_COLOR'):
    Colors.disable()


def print_status(msg, status="info"):
    """Print a status message with icon and color."""
    icons = {"info": "i", "success": "+", "warning": "!", "error": "x"}
    colors = {
        "info": Colors.CYAN,
        "success": Colors.GREEN,
        "warning": Colors.YELLOW,
        "error": Colors.RED
    }
    color = colors.get(status, Colors.RESET)
    icon = icons.get(status, "")
    print(f"{color}[{icon}] {msg}{Colors.RESET}")


def print_header(text):
    """Print a header with separator lines."""
    print(f"\n{Colors.BOLD}{Colors.CYAN}{'='*50}{Colors.RESET}")
    print(f"{Colors.BOLD}{Colors.CYAN}{text}{Colors.RESET}")
    print(f"{Colors.BOLD}{Colors.CYAN}{'='*50}{Colors.RESET}\n")


def print_thinking(msg):
    """Print a thinking/processing message."""
    print(f"{Colors.DIM}... {msg}{Colors.RESET}", flush=True)


def print_code(code, language="vba"):
    """Print a code block with separator lines."""
    print(f"\n{Colors.BOLD}--- Generated {language.upper()} ---{Colors.RESET}")
    print(f"{Colors.DIM}{'-'*50}{Colors.RESET}")
    print(code)
    print(f"{Colors.DIM}{'-'*50}{Colors.RESET}\n")


def spinner(message, duration=2):
    """Simple spinner for indeterminate progress."""
    chars = ["|", "/", "-", "\\"]
    start = time.time()
    i = 0
    while time.time() - start < duration:
        print(f"\r{Colors.CYAN}{chars[i % len(chars)]} {message}{Colors.RESET}",
              end="", flush=True)
        time.sleep(0.1)
        i += 1
    print(f"\r{' '*50}", end="", flush=True)


def format_table(data, headers=None):
    """Format 2D array as ASCII table."""
    if not data:
        return "(empty)"

    # Calculate column widths
    all_rows = data if headers is None else [headers] + data
    col_widths = []
    for i in range(max(len(row) for row in all_rows)):
        max_w = 0
        for row in all_rows:
            if i < len(row):
                max_w = max(max_w, len(str(row[i])))
        col_widths.append(max_w)

    lines = []

    # Header
    if headers:
        header_line = " | ".join(
            str(h).ljust(col_widths[i]) for i, h in enumerate(headers)
        )
        lines.append(f"{Colors.BOLD}{header_line}{Colors.RESET}")
        lines.append("-+-".join("-" * w for w in col_widths))

    # Data rows
    for row in data:
        line = " | ".join(
            str(row[i]).ljust(col_widths[i]) if i < len(row) else " " * col_widths[i]
            for i in range(len(col_widths))
        )
        lines.append(line)

    return "\n".join(lines)


def prompt_input(prompt_text="excel> "):
    """Get user input with styled prompt."""
    try:
        return input(f"{Colors.CYAN}{prompt_text}{Colors.RESET}").strip()
    except (EOFError, KeyboardInterrupt):
        print()
        return None
