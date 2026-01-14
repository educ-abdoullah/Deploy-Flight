#!/bin/bash
playwright install chromium
gunicorn app:app --bind 0.0.0.0:10000
