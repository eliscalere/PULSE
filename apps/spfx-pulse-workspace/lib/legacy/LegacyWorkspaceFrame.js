import * as React from 'react';
import { legacyWorkspaceHtml } from './LegacyWorkspaceHtml';
var MIN_HEIGHT = 1200;
export function LegacyWorkspaceFrame() {
    var iframeRef = React.useRef(null);
    var _a = React.useState(MIN_HEIGHT), height = _a[0], setHeight = _a[1];
    React.useEffect(function () {
        function onMessage(event) {
            if (!event.data || event.data.source !== 'aewttr-legacy-frame') {
                return;
            }
            var nextHeight = Number(event.data.height);
            if (!Number.isNaN(nextHeight) && nextHeight > 0) {
                setHeight(Math.max(MIN_HEIGHT, nextHeight + 24));
            }
        }
        window.addEventListener('message', onMessage);
        return function () { return window.removeEventListener('message', onMessage); };
    }, []);
    return (React.createElement("iframe", { ref: iframeRef, title: "AEWTTR-PULSE", srcDoc: legacyWorkspaceHtml, style: {
            width: '100%',
            height: "".concat(height, "px"),
            border: '0',
            display: 'block',
            backgroundColor: '#eef0f8'
        } }));
}
//# sourceMappingURL=LegacyWorkspaceFrame.js.map