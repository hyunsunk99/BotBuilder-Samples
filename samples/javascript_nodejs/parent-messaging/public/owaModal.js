microsoftTeams.initialize();
var exports = {};

exports.run = function () {
    // After successful generation of an email draft from the conversation, save the draft id in local storage and then redirect to OWA compose box.
    // If user discards the draft from client, client will use the stored draft id and delete that email
    
    var urlParams = new URLSearchParams(window.location.search);
    var draftId = urlParams.get('draftId'); // already encoded
    console.log(draftId);

    var outlookOrigin = 'https://outlook.office.com';
    var owaFrameId = 'sto-owa-frame';
    var body = document.getElementsByTagName('body')[0];

    // host a full iframe set to owa instead and proxy post message calls
    // between parent (Teams) and child (OWA)
    var src = `${outlookOrigin}/mail/opxdeeplink/compose/${draftId}?isanonymous=true&opxAuth&hostApp=teams&useOwaTheme`;
    if (desktopVersion) {
        src = `${src}&desktopVersion=${encodeURIComponent(desktopVersion)}`;
    }
    if (owaParams) {
        src = `${src}&${owaParams}`;
    }
    body.innerHTML = `<iframe id='${owaFrameId}' src='${src}'></iframe>`;

    var owaFrame = document.getElementById(owaFrameId).contentWindow;

    window.addEventListener('message', function receiveMessage(event) {
        if (!event) { return; }

        var messageOrigin = event.origin || event.originalEvent.origin

        if (messageOrigin === window.location.origin) {
        // post message to OWA from Teams
        owaFrame.postMessage(event.data, outlookOrigin);
        } else {
        // post message to Teams from OWA
        window.parent.postMessage(event.data, window.location.origin);
        }
    });
}

return exports;