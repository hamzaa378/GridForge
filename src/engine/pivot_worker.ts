import {runPivotJob,runPivotJobs,type PivotBatchJobRequest,type PivotJobRequest} from './pivot_job.js';

const workerScope=self as unknown as {
  onmessage:((event:MessageEvent<PivotJobRequest|PivotBatchJobRequest>)=>void)|null;
  postMessage:(message:ReturnType<typeof runPivotJob>|ReturnType<typeof runPivotJobs>)=>void;
};

workerScope.onmessage=event=>workerScope.postMessage('specs' in event.data?runPivotJobs(event.data):runPivotJob(event.data));
