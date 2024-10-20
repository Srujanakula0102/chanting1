const CLIENT_ID = "771591691116-0v0o5n4dl0dpmt0cjnped1e23d76egle.apps.googleusercontent.com" ;
const API_KEY = "AIzaSyD3XkSg8I29ug-if7RenmiDxNi69jcwLtk" ;
const SCOPES = 'https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/drive.readonly';
const DISCOVERY_DOCS = ["https://www.googleapis.com/discovery/v1/apis/drive/v3/rest"];

let pickerApiLoaded = false;
let oauthToken;

function onApiLoad() {
    gapi.load('auth', {'callback': onAuthApiLoad});
    gapi.load('picker', {'callback': onPickerApiLoad});
}

function onAuthApiLoad() {
    gapi.auth.authorize(
        {
            'client_id': CLIENT_ID,
            'scope': SCOPES,
            'immediate': false
        },
        handleAuthResult);
}

function onPickerApiLoad() {
    pickerApiLoaded = true;
    document.getElementById('pick').onclick = createPicker;
}

function handleAuthResult(authResult) {
    if (authResult && !authResult.error) {
        oauthToken = authResult.access_token;
        createPicker();
    }
}

function createPicker() {
    if (pickerApiLoaded && oauthToken) {
        const view = new google.picker.View(google.picker.ViewId.DOCS);
        view.setMimeTypes("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        const picker = new google.picker.PickerBuilder()
            .enableFeature(google.picker.Feature.MULTISELECT_ENABLED)
            .setOAuthToken(oauthToken)
            .addView(view)
            .setDeveloperKey(API_KEY)
            .setCallback(pickerCallback)
            .build();
        picker.setVisible(true);
    }
}

function pickerCallback(data) {
    if (data.action === google.picker.Action.PICKED) {
        const files = data.docs;
        const promises = files.map(file => gapi.client.drive.files.get({
            fileId: file.id,
            alt: 'media'
        }).then(response => {
            const binary = atob(response.body);
            const workbook = XLSX.read(binary, {type: 'binary'});
            return XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], {header: 1});
        }));

        Promise.all(promises).then(sheets => {
            const mergedData = [].concat(...sheets);
            const mergedSheet = XLSX.utils.aoa_to_sheet(mergedData);
            const newWorkbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(newWorkbook, mergedSheet, 'Merged Data');
            const wbout = XLSX.write(newWorkbook, {bookType: 'xlsx', type: 'binary'});

            const blob = new Blob([s2ab(wbout)], {type: 'application/octet-stream'});
            const url = URL.createObjectURL(blob);
            const a = document.getElementById('downloadLink');
            a.href = url;
            a.download = 'merged.xlsx';
            a.style.display = 'block';
            a.textContent = 'Download Merged File';
        });
    }
}

function s2ab(s) {
    const buf = new ArrayBuffer(s.length);
    const view = new Uint8Array(buf);
    for (let i = 0; i < s.length; i++) {
        view[i] = s.charCodeAt(i) & 0xFF;
    }
    return buf;
}

gapi.load('client:auth2', onApiLoad);