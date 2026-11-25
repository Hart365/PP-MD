import type { ParsedSolution } from '../types/solution';
import { parseSolutionZip } from '../parser/solutionParser';
import SolutionWorker from './solution.worker?worker&inline';

export const PARSE_CONCURRENCY = 2;

export interface ParseSettledCallbacks {
  onProgress: (index: number, percent: number) => void;
  onSettled: (index: number, result: PromiseSettledResult<ParsedSolution>) => void;
  onFallback?: (message: string) => void;
}

class WorkerUnavailableError extends Error {}

function parseInWorker(file: File, signal: AbortSignal, onProgress: (percent: number) => void): Promise<ParsedSolution> {
  return new Promise((resolve, reject) => {
    let worker: Worker;
    try {
      worker = new SolutionWorker();
    } catch (error) {
      reject(new WorkerUnavailableError(error instanceof Error ? error.message : 'The worker could not be created.'));
      return;
    }
    const finish = (error?: Error, solution?: ParsedSolution) => {
      signal.removeEventListener('abort', abort);
      worker.terminate();
      if (error) reject(error);
      else resolve(solution!);
    };
    const abort = () => finish(new DOMException('Solution processing was cancelled.', 'AbortError'));
    signal.addEventListener('abort', abort, { once: true });
    worker.onmessage = (event: MessageEvent<
      | { kind: 'progress'; percent: number }
      | { kind: 'result'; solution: ParsedSolution }
      | { kind: 'error'; name: string; message: string }
    >) => {
      if (event.data.kind === 'progress') onProgress(event.data.percent);
      else if (event.data.kind === 'result') finish(undefined, event.data.solution);
      else {
        const error = new Error(event.data.message);
        error.name = event.data.name;
        finish(error);
      }
    };
    worker.onerror = (event) => {
      event.preventDefault();
      finish(new WorkerUnavailableError(event.message || 'The solution processing worker could not run.'));
    };
    worker.onmessageerror = () => finish(new WorkerUnavailableError('The solution processing worker returned unreadable data.'));
    void file.arrayBuffer().then((data) => worker.postMessage({ data }, [data]), (error) => finish(error));
  });
}

export async function parseSolutionsConcurrently(
  files: File[],
  signal: AbortSignal,
  callbacks: ParseSettledCallbacks,
): Promise<PromiseSettledResult<ParsedSolution>[]> {
  const outcomes: PromiseSettledResult<ParsedSolution>[] = new Array(files.length);
  let nextIndex = 0;
  let compatibilityMode = false;
  const parseOne = async (index: number) => {
    let outcome: PromiseSettledResult<ParsedSolution>;
    try {
      signal.throwIfAborted();
      let solution: ParsedSolution;
      if (!compatibilityMode) {
        try {
          solution = await parseInWorker(files[index], signal, (percent) => callbacks.onProgress(index, percent));
        } catch (error) {
          if (!(error instanceof WorkerUnavailableError)) throw error;
          compatibilityMode = true;
          callbacks.onFallback?.('Background workers are unavailable in this host. Continuing solution parsing in compatibility mode.');
          solution = await parseSolutionZip(files[index], (percent) => callbacks.onProgress(index, percent), signal);
        }
      } else {
        solution = await parseSolutionZip(files[index], (percent) => callbacks.onProgress(index, percent), signal);
      }
      outcome = { status: 'fulfilled', value: solution };
    } catch (reason) {
      outcome = { status: 'rejected', reason };
    }
    outcomes[index] = outcome;
    callbacks.onSettled(index, outcome);
  };
  await Promise.all(Array.from({ length: Math.min(PARSE_CONCURRENCY, files.length) }, async () => {
    while (nextIndex < files.length && !signal.aborted) await parseOne(nextIndex++);
  }));
  return outcomes;
}
