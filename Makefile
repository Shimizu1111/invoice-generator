.PHONY: open install deploy cli-estimate cli-invoice cli-invoice-from-estimate cli-estimate-from-git help

help: ## Show this help
	@grep -E '^[a-zA-Z_-]+:.*?## .*$$' $(MAKEFILE_LIST) | awk 'BEGIN {FS = ":.*?## "}; {printf "  \033[36m%-28s\033[0m %s\n", $$1, $$2}'

open: ## Open the web UI in browser (via local server)
	@echo "http://localhost:8000 で起動中... Ctrl+C で停止"
	@open http://localhost:8000 &
	cd docs && python3 -m http.server 8000

deploy: ## Deploy to GitHub Pages (commit & push docs/)
	git add docs/
	git commit -m "Deploy to GitHub Pages" || echo "No changes to commit"
	git push origin main

install: ## Install dependencies
	npm install

cli-estimate: ## Create estimate via CLI (usage: make cli-estimate JSON='{"clientName":"...","subject":"...","items":[...]}')
	node index.js estimate '$(JSON)'

cli-invoice: ## Create invoice via CLI (usage: make cli-invoice JSON='{"clientName":"...","subject":"...","items":[...]}')
	node index.js invoice '$(JSON)'

cli-invoice-from-estimate: ## Create invoice from estimate (usage: make cli-invoice-from-estimate ID=<spreadsheet-id>)
	node index.js invoice-from-estimate $(ID)

cli-estimate-from-git: ## Create estimate from git history (usage: make cli-estimate-from-git JSON='{"repoPath":"...","clientName":"...","subject":"...","since":"..."}')
	node index.js estimate-from-git '$(JSON)'
