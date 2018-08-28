#target indesign  
myDoc = app.activeDocument;  
var textArr = new Array(), mRes, myPit, count, done;  
mFramesToMove = new Array();  
myPit = myDoc.allPageItems;    // array with doc pageItems (masterpages and groups included)  
for (var k = 0; k < myPit.length; k++)  
    if (myPit[k].constructor.name == "TextFrame" && myPit[k].overflows )  
    textArr.push(myPit[k]);    // array with overset textFrames  
count = textArr.length;  
  
if (count) mRes = mDialog (count);  
else {alert ("No overset texts found"); exit (); }  
  
for (var i = 0; i < textArr.length; i++) {  
    done = mShrink (textArr[i], mRes);  
    if (!done[0] ) {  
        alert (done[1] + "\rFill a dialog field with a number");   
        exit; }  
    }  
  
    if (mFramesToMove.length) {  
        if (myDoc.layers.item("Overset frames").isValid)  
            mLayer = myDoc.layers.item("Overset frames");  
        else   
            mLayer = myDoc.layers.add({LocationOptions: LocationOptions.AT_BEGINNING, name: "Overset frames"});  
        for (var i = 0; i < mFramesToMove.length; i++)  
        mFramesToMove[i].itemLayer = mLayer;  
        alert ("Some frames (" + mFramesToMove.length + ") stay overset\rTake a look at layer 'Overset Frames' ");  
        }  
    else alert (count + " frames overset fixed");  
  
function mDialog (number) {// "What to do" dialog  
    var w = new Window("dialog","What to do?", undefined, {closeButton: false});  
    w.add ("statictext", undefined, "Number of overset texts found:  " + number.toString() );  
  
    var myTop = w.add ("panel");  
    myTop.alignment = "left";  
    myTop.orientation = "row";  
  
    var myCheckBx = myTop.add ("group");  
    myCheckBx.orientation = "column";  
    myCheckBx.alignChildren = "left";  
    var Ch1 = myCheckBx.add("checkbox", undefined, "point size limit: ");  
    var Ch2 = myCheckBx.add("checkbox", undefined, "hori scale limit: ");  
  
    var myEditBx = myTop.add ("group");  
    myEditBx.orientation = "column";  
    var Ed1 = myEditBx.add("edittext", undefined, undefined);  
    Ed1.enabled = false; Ed1.characters = 4;  
    var Ed2 = myEditBx.add("edittext", undefined, undefined);  
    Ed2.enabled = false; Ed2.characters = 4;  
  
    Ch1.onClick = function () {   
        if (Ch1.value) Ed1.enabled = true;  
        else Ed1.enabled = false;}  
    Ch2.onClick = function () {   
        if (Ch2.value) Ed2.enabled = true;  
        else Ed2.enabled = false;}  
  
    var myStatBx = myTop.add ("group");  
    myStatBx.orientation = "column";  
    myStatBx.add("statictext", undefined, " pt");  
    myStatBx.add("statictext", undefined, " %");  
  
    var b = w.add ("group");  
        b.add ("button", undefined, "OK", {name: "ok"});  
        b.add ("button", undefined, "Anuluj", {name: "cancel"});  
  
    if (w.show() == 1)   
        return [Ch1.value, Ed1.text, Ch2.value, Ed2.text];  
    else exit();  
    }  
  
function mShrink (textFr, ToDoArr) {    // input textFrame and array with mDialog result  
    var pSizeLimAc, pSizeLim, pSizeCurr, pSizeDiff, pSizeStepN, pSizeStep,  
        hScaleAc, hScaleLim, hScaleCurr, hScaleDiff, hScaleStepN, hScaleStep,  
        ToDo;  
    pSizeLimAc = ToDoArr[0];  
    hScaleAc = ToDoArr[2];  
  
    pSizeLim = ToDoArr[1];  
    if (pSizeLimAc) {  
        if (isNaN (parseFloat(pSizeLim) ) ) return [false, "point size is not a number"];  
        else pSizeLim = parseFloat(pSizeLim);  
        pSizeCurr = textFr.parentStory.pointSize;  
        pSizeDiff = pSizeCurr - pSizeLim;   
        pSizeStepN = pSizeDiff / 0.1;  
        pSizeStep = 0.1;  
        }  
  
    hScaleLim = ToDoArr[3];  
    if (hScaleAc) {  
        if (isNaN (parseFloat(hScaleLim) ) ) return [false, "hori scale is not a number"];  
        else hScaleLim = parseFloat(hScaleLim);  
        hScaleCurr = textFr.parentStory.horizontalScale;  
        hScaleDiff = hScaleCurr - hScaleLim;   
        if (pSizeLimAc) hScaleStepN = pSizeStepN;   
        else hScaleStepN = hScaleDiff;  
        hScaleStep = Math.round(hScaleDiff/hScaleStepN*10)/10;  
        }  
    if (!(hScaleDiff > 0) ) hScaleAc = false;  
    if (!(pSizeDiff > 0) ) pSizeLimAc = false;  
    if (pSizeLimAc && hScaleAc) ToDo = 3;  
    else if (pSizeLimAc) ToDo = 1;  
        else if (hScaleAc) ToDo = 2;  
            else ToDo = 0;  
  
    switch (ToDo) {  
        case 1: {   
            while (pSizeStepN-- && textFr.overflows)  
                textFr.parentStory.pointSize -= pSizeStep;  
                break; }  
        case 2: {   
            while (hScaleStepN-- && textFr.overflows)  
                textFr.parentStory.horizontalScale -= hScaleStep;  
                break; }  
        case 3: {   
            while (hScaleStepN-- && textFr.overflows)  
                {  
                textFr.parentStory.pointSize -= pSizeStep;  
                textFr.parentStory.horizontalScale -= hScaleStep;  
                }  
                break; }  
        default: break;  
        }  
    if (textFr.overflows) mFramesToMove.push(textFr);  
    return [true];  
    }  