/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global  Office, Word */
async function tryCatch(callback) {
  try {
    return callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    setError(error);
  }
}
function getFile(){
    return new Promise((resolve, reject) => {
      Office.context.document.getFileAsync(Office.FileType.Text, { sliceSize: 4194304  /*64 KB*/ },
        function (result) {
            if (result.status == "succeeded") {
                // If the getFileAsync call succeeded, then
                // result.value will return a valid File Object.
                var myFile = result.value;
                var sliceCount = myFile.sliceCount;
                var slicesReceived = 0, gotAllSlices = true, docdataSlices = [];

                // Get the file slices.
                getSliceAsync(myFile, 0, sliceCount, gotAllSlices, docdataSlices, slicesReceived)
                  .then((fileData) => {
                    resolve(fileData);
                  }).catch(error => {
                    setError(error); reject(error)
                  });
            }
            else {
                setError(result.error.message);
                reject(result.error.message);
            }
        });
      });
}


function getSliceAsync(file, nextSlice, sliceCount, gotAllSlices, docdataSlices, slicesReceived) {
    return new Promise((resolve, reject) => {
      file.getSliceAsync(nextSlice, function (sliceResult) {
          if (sliceResult.status == "succeeded") {
              if (!gotAllSlices) { // Failed to get all slices, no need to continue.                                });
                  return;                                                                                           
              }                                                                                                     
                                                                                                                     
              // Got one slice, store it in a temporary array.                                                      
              // (Or you can do something else, such as                                                             
              // send it to a third-party server.)
              docdataSlices[sliceResult.value.index] = sliceResult.value.data;                                      
              if (++slicesReceived == sliceCount) {                                                                 
                  // All slices have been received.                                                                 
                  file.closeAsync();                                                                                
                  resolve(docdataSlices); 
              }                                                                                                     
              else {                                                                                                
                  getSliceAsync(file, ++nextSlice, sliceCount, gotAllSlices, docdataSlices, slicesReceived);        
              }                                                                                                     
          }                                                                                                         
          else {                                                                                                    
              gotAllSlices = false;                                                                                 
              file.closeAsync();                                                                                    
              setError(sliceResult.error.message);                                                                  
              reject(sliceResult.error.message);
          }                                                                                                         
      });
    });
}

 function setError(message) { 
    const errorEl = document.getElementById('error');
    errorEl.style.display = '';
    errorEl.innerText = message;
 }

 function clearError() { 
    const errorEl = document.getElementById('error');
    errorEl.innerText = '';
    errorEl.style.display = 'none';
 }

function getChunkSize() {
  // Gets Chunk Size  
  const chunkSizeEl = document.getElementById('chunk-size');
  let chunkSize = 4000; 
  const chunkSizeValue = chunkSizeEl.value;
  if (chunkSizeValue.length > 0) {
    const value = parseInt(chunkSizeValue, 10);
    const minValue = 500; 
    if (isNaN(value)) {
      setError('Error: Section character count value must be a valid number');
    } else if (value < minValue) {
      setError(`Error: Section character count must be greater than ${minValue}`);
    } else
      chunkSize = value;
    }
  return chunkSize;
}

function setSelectedValue(selectionEl, valueToSet) {
    for (var i = 0; i < selectionEl.options.length; i++) {
        if (selectionEl.options[i].text== valueToSet) {
            selectionEl.options[i].selected = true;
            return;
        }
    }
}

function getSelectionOptions(chunkCount) {
  const selectionEl = document.getElementById('chunk-number');
  const previousValue = selectionEl.value;

  while (selectionEl.options.length > 0) {                
    selectionEl.remove(0);
  } 

  for (let i = 1; i <= chunkCount; i++) {
    const option = document.createElement("option"); 
    option.text = i;
    option.value = i;
    selectionEl.add(option);
  }

  setSelectedValue(selectionEl, previousValue);
}


 function getData() {
    startLoading();
    return new Promise((resolve, reject) => {
      return Word.run(async function () {
          return getFile() 
            .then(function(result){
               let bodyText = '';
               result.forEach((paragraph, index) => {
                 if ((paragraph || '').length > 0) {
                   if (index === result.length) {
                     bodyText += paragraph;  
                     return;
                   }
                   
                   bodyText += `${paragraph}`
                 }
               });
               const wordsArray = bodyText.split(/[\s\n]+/).filter(w => w.length > 0);

               const wordCount = wordsArray.length;
               const wordCountEl = document.getElementById('word-count');
               wordCountEl.innerText = wordCount;

               clearError();
               const charaCountEl = document.getElementById('character-count');
               charaCountEl.innerText = bodyText.length;

               stopLoading();
               resolve({ wordsArray, bodyText });

             }).catch(error => {
               stopLoading();
               setError(error);
               reject(error);
             });
           }).catch((error) => {
               setError(error);
               reject(error);
           });
    });
}

function updateUI() {
  getData().then((data) => {
    const { bodyText } = data; 
    const chunkSize = getChunkSize();
    const chunkCountEl = document.getElementById('chunk-count');
    const chunkCount = Math.ceil(bodyText.length / chunkSize);
    chunkCountEl.innerText = chunkCount;
    
    const selectionEl = document.getElementById('chunk-number');
    
    if (document.activeElement !== selectionEl) {
      getSelectionOptions(chunkCount);
    }

  });
}


function copySectionToClipboard(bodyText) {
  const selectionEl = document.getElementById('chunk-number');
  const selectionValue = selectionEl.value;
  if (selectionValue) {
    const chunkNumber = parseInt(selectionValue, 10);

    const chunkSize = getChunkSize();
    const chunkStart = (chunkNumber - 1) * chunkSize;
    const chunkEnd = chunkStart + chunkSize;

    const sectionText = bodyText.substring(chunkStart, chunkEnd);
    const allWords = bodyText.split(/[\s\n]+/).filter(w => w.length > 0)

    const wordsArray = sectionText.split(/[\s\n]+/).filter(w => w.length > 0);

    if (typeof wordsArray[1] !== 'undefined') {
      const firstWordIndex = allWords.indexOf(wordsArray[1]) - 1;
      const firstWord = allWords[firstWordIndex];

      if (typeof firstWord !== 'undefined') {
        wordsArray[0] = firstWord;
      }
    }

    let textChunk = '';
    for(let i = 0; i < wordsArray.length - 1; i++) {
      if (typeof wordsArray[i] !== 'undefined') {
        const chunkAddition = `${textChunk} ${wordsArray[i + 1]}`;

        if (i === wordsArray.length - 2 || typeof wordsArray[i + 1] === 'undefined' || chunkAddition.length > chunkSize) {
          textChunk += wordsArray[i];
          break;
        }
        textChunk += `${wordsArray[i]} `;
      } 
    }
    
    // Use the 'out of viewport hidden text area' trick
    const textArea = document.getElementById("copy-to-clip");
    textArea.value = textChunk;
        
    // Move textarea out of the viewport so it's not visible
    textArea.style.position = "absolute";
    textArea.style.left = "-999999px";
  
    textArea.select();
  
    try {
        document.execCommand('copy');
        showCopiedBox();
    } catch (error) {
        console.error(error);
    } 
  }
}

function showCopiedBox() {
  const copiedBox = document.getElementById("copied-box");

  const selectionEl = document.getElementById('chunk-number');
  const copiedInfo = document.getElementById('copied-info');
  copiedInfo.innerText = `Section ${selectionEl.value} copied to clipboard`;

  copiedBox.style.opacity = 1;
  copiedBox.style.display = '';

  const waitTime = 1200;
  setTimeout(() => {
    copiedBox.style.opacity = 0;
  }, waitTime);

  setTimeout(() => {
    copiedBox.style.display = 'none';
  }, waitTime + 500);

}

function startLoading() {
  const button = document.getElementById("run");
  button.style.opacity = 0.7;
  button.style.pointerEvents = 'none';
}

function stopLoading() {
  const button = document.getElementById("run");
  button.style.opacity = 1;
  button.style.pointerEvents = '';
}

function debounce(func, timeout = 1000){
  let timer;
  return (...args) => {
    clearTimeout(timer);
    timer = setTimeout(() => { func.apply(this, args); }, timeout);
  };
}


Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
     updateUI();
     document.getElementById('sideload-msg').style.display = 'none';
     document.getElementById('app-body').style.display = 'flex';
     document.getElementById('run').onclick = () => tryCatch(run);
     document.body.onclick = () => tryCatch(updateUI);
     document.getElementById('chunk-size').addEventListener('input', () => debounce(tryCatch(updateUI)));
  }
});


export async function run() {
  return Word.run(async (context) => {
    getData()
      .then((data) =>  {
        const { bodyText } = data;
        copySectionToClipboard(bodyText)
      });

    await context.sync();
  });
}
