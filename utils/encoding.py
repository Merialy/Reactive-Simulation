import sys
import io

def setup_utf8_output() -> None:
    """
    Налаштування UTF-8 кодування для коректного виводу кириличних символів у консолі.
    """
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")
