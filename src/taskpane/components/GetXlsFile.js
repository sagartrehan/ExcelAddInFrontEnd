
// The following example gets the document in Office Open XML ("compressed") format in 65536 bytes (64 KB) slices.
// Note: The implementation of app.showNotification in this example is from the Visual Studio template for Office Add-ins.
export function getDocumentAsCompressed(callbacks) {
    Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 65536 /*64 KB*/ },
        function (result) {
            if (result.status == "succeeded") {
                // If the getFileAsync call succeeded, then
                // result.value will return a valid File Object.
                var myFile = result.value;
                var sliceCount = myFile.sliceCount;
                var slicesReceived = 0, gotAllSlices = true, docdataSlices = [];
                console.log("File size:" + myFile.size + " #Slices: " + sliceCount);
                // Get the file slices.
                getSliceAsync(myFile, 0, sliceCount, gotAllSlices, docdataSlices, slicesReceived, callbacks);
            }
            else {
                console.log("Error fetching file asyncronously..");
            }
        });
}

function getSliceAsync(file, nextSlice, sliceCount, gotAllSlices, docdataSlices, slicesReceived, callbacks) {
    file.getSliceAsync(nextSlice, function (sliceResult) {
        if (sliceResult.status == "succeeded") {
            if (!gotAllSlices) { // Failed to get all slices, no need to continue.
                return;
            }

            // Got one slice, store it in a temporary array.
            // (Or you can do something else, such as
            // send it to a third-party server.)
            docdataSlices[sliceResult.value.index] = sliceResult.value.data;
            if (++slicesReceived == sliceCount) {
                // All slices have been received.
                file.closeAsync();
                onGotAllSlices(docdataSlices, callbacks);
            }
            else {
                getSliceAsync(file, ++nextSlice, sliceCount, gotAllSlices, docdataSlices, slicesReceived, callbacks);
            }
        }
        else {
            gotAllSlices = false;
            file.closeAsync();
            app.showNotification("getSliceAsync Error:", sliceResult.error.message);
        }
    });
}

function onGotAllSlices(docdataSlices, callbacks) {
    var docdata = [];
    for (var i = 0; i < docdataSlices.length; i++) {
        docdata = docdata.concat(docdataSlices[i]);
    }

    const file = new File([new Uint8Array(docdata)], 'testfile.xls', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    uploadFile(file, callbacks)
}


function uploadFile(file, callbacks) {
    const data = new FormData()
    data.append("file", file);
    data.append('data', JSON.stringify({ "id": "1" }))
    axios
        .post("127.0.0.1:9001/v1/excelRead", data, {            
        })
        .then(res => { // then print response status
            console.log(res.statusText)
            if (res.responseCode != 0) {
                callbacks("error processing your file", null)
            } else {
                callbacks(null, res.processedFilePath)
            }
        })
}