# Results

`eval_gt.json` is the reference score file for the bundled ground-truth
answers. Use it to check that the evaluator is working on your machine:

```bash
python eval/eval.py gt --output results/eval_gt_check.json
```

The expected `overall_avg_score_rate` is `95.5` across 200 scored tasks.
Raw evaluator failure messages are omitted from the checked-in result files;
the per-task `num_errors` fields are retained.

`eval_gpt-5.5.json` is one representative model result file. It records
the evaluation summary for our gpt-5.5 single-turn reproduction run:

- `total_tasks`: 200
- `total_scored`: 200
- `overall_avg_score_rate`: 37.9

The paper reports GPT-5.5 average SR `36.1`; this reproduced single-run score
is close enough to use as a model-score sanity reference. The raw GPT-5.5
output files are not included here, so this JSON is for comparison rather than
for re-running the model-output score directly from this release.
