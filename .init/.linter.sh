#!/bin/bash
cd /home/kavia/workspace/code-generation/google-sheets-automation-suite-4168-4177/sheets_script_frontend
npm run build
EXIT_CODE=$?
if [ $EXIT_CODE -ne 0 ]; then
   exit 1
fi

