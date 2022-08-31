Office.initialize = function () {
}

function validateBody(event) {
    Office.onReady().then(function() {
        var contextInfo = Office.context.diagnostics;
        //Execute the add-in logic only if it is Outlook application running on Windows
        if(contextInfo.platform == 'PC'){
            event.completed({ allowEvent: true });
        }
    });
}


if (typeof exports !== 'undefined') {
    exports.validateBody = validateBody;
}

