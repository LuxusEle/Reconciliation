import { extract, reconcile, exportReport } from './reconciliation.js';
self.onmessage = (event) => {
    try {
        const dn = extract(event.data.dn, 'dn');
        const core = extract(event.data.core, 'core');
        const report = reconcile(dn, core, event.data.direction);
        const file = exportReport(report, core, dn);
        self.postMessage({ report, file }, { transfer: [file] });
    }
    catch (error) {
        self.postMessage({ error: error instanceof Error ? error.message : 'Unable to process these workbooks.' });
    }
};
