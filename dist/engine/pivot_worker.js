import { runPivotJob, runPivotJobs } from './pivot_job.js';
const workerScope = self;
workerScope.onmessage = event => workerScope.postMessage('specs' in event.data ? runPivotJobs(event.data) : runPivotJob(event.data));
