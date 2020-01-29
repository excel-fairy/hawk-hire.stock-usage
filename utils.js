/**
 * Converts a GDrive folder URL to an ID
 * @param url the URl
 * @returns {*} The ID as a String
 */
function folderUrlToId(url) {
    var regex = /^https:\/\/drive\.google\.com\/drive\/folders\/(.*)$/g;
    var match = regex.exec(url);
    return match[1];
}

/**
 * Converts a Google spreadsheet URL to an ID
 * @param url the URl
 * @returns {*} The ID as a String
 */
function spreadsheetUrlToId(url) {
    var regex = /^https:\/\/docs\.google\.com\/spreadsheets\/d\/(.*?)(\/edit#gid=.*)?$/g;
    var match = regex.exec(url);
    return match[1];
}
