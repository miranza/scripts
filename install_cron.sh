#!/bin/sh
set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PYTHON="$(pyenv which python)"

echo "Project dir : $SCRIPT_DIR"
echo "Python : $PYTHON"

MOVCONT="0 5 * * * cd $SCRIPT_DIR && $PYTHON movcont_v1.py"
TRIAL="0 6 * * * cd $SCRIPT_DIR && $PYTHON trial_balances_v1.py"

# Remove old entries for these scripts if they exist
CURRENT=$(crontab -l 2>/dev/null | grep -v "movcont_v1.py" | grep -v "trial_balances_v1.py")

echo "$CURRENT
$MOVCONT
$TRIAL" | crontab -

echo "Crontab instalado:"
crontab -l | grep -E "movcont_v1|trial_balances_v1"
