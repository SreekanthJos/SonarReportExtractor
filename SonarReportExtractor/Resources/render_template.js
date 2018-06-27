var page = require('webpage').create();
page.open('[UserURL]', function () {
    [RENDER_CODE]
    phantom.exit();
});