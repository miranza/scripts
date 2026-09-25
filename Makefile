.PHONY: format check install clean

PIP    = pip
SRC    = ./

install:
	$(PIP) install isort black

format:
	isort --profile black -l 100 $(SRC)
	black -l 100 $(SRC)

check:
	isort --profile black -l 100 $(SRC) --check
	black -l 100 $(SRC) --check

movcont:
	python movcont_v1.py update_last_month --acc-from 600000  --acc-to 699999

trial:
	python trial_balances_v1.py --auto-update

companies:
	python get_companies.py

clean:
	find data -maxdepth 1 -name '*backup_*' -mtime +7 -delete
	rm -f logs/*.log
