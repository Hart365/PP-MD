/// <reference lib="webworker" />

import { parseSolutionZip } from '../parser/solutionParser';

self.onmessage = async (event: MessageEvent<{ data: ArrayBuffer }>) => {
  try {
    const solution = await parseSolutionZip(new Blob([event.data.data]), (percent) => {
      self.postMessage({ kind: 'progress', percent });
    });
    self.postMessage({ kind: 'result', solution });
  } catch (error) {
    self.postMessage({
      kind: 'error',
      name: error instanceof Error ? error.name : 'Error',
      message: error instanceof Error ? error.message : 'Solution parsing failed.',
    });
  }
};
