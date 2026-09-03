# -*- coding: utf-8 -*-
"""Точка входа: отчёт по товарам для участия в акции Ozon.

Запуск:
    python main.py
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from ui.app import run

if __name__ == "__main__":
    run()
