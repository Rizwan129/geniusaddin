export async function getValue(key) {
    let value = await OfficeRuntime.storage.getItem(key);
    return value;
}

export async function saveProp(key, value) {
    OfficeRuntime.storage.setItem(key, value);
}

export async function removeProp(key) {
    OfficeRuntime.storage.removeItem(key);
}

export async function getCategories() {
    var GENIUSSHEETS_URL = 'https://geniussheets.herokuapp.com';
    var access_token = await getValue('access_token');
    let url = GENIUSSHEETS_URL + '/api/getCategories/';
    var headers = {
        'Authorization': 'Token ' + access_token
    };
    var options = {
        'method': 'GET',
        'contentType': 'application/json',
        'headers': headers,
        'muteHttpExceptions': true
    };

    const response = await fetch(url, options);
    if (response.ok) {
        const jsonResponse = await response.json();
        return jsonResponse.data;
    } else if (response.getResponseCode() !== 200) {
        throw new Error('Couldnot fetch categories. Check Backend.');
    }
}
