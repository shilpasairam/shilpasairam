#!/bin/sh

# sanity = ['C37915', 'C29642']
# smoketest = ['smoketest', 'smoketest_nma']

sanity="C37915 C29642"

for tc in $arg1 
do
  python -m pytest -v -m tc --self-contained-html --nunitxml=Reports/test-results.xml .\testCases\ --env=$arg2
done