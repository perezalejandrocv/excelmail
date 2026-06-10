name: ExcelMailSync

on:
  schedule:
    - cron: "0 * * * *"   # cada hora
  workflow_dispatch:      # ejecución manual opcional

jobs:
  run-script:
    runs-on: ubuntu-latest

    steps:
      - name: Checkout repo
        uses: actions/checkout@v4

      - name: Setup Python
        uses: actions/setup-python@v5
        with:
          python-version: "3.11"

      - name: Install dependencies
        run: pip install msal requests

      - name: Run script
        env:
          CLIENT_ID: ${{ secrets.CLIENT_ID }}
          TENANT_ID: ${{ secrets.TENANT_ID }}
          CLIENT_SECRET: ${{ secrets.CLIENT_SECRET }}
        run: python ExcelMailSync.py
