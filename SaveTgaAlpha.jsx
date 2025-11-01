#target photoshop
app.bringToFront();

// 獲取上次保存的路徑
function getLastSavePath() {
    var file = new File(Folder.temp + "/lastTGASavePath.txt");
    if (file.exists) {
        file.open('r');
        var path = file.read();
        file.close();
        return path;
    }
    return "";
}

// 設置上次保存的路徑
function setLastSavePath(path) {
    var file = new File(Folder.temp + "/lastTGASavePath.txt");
    file.open('w');
    file.write(path);
    file.close();
}

function saveAsTGAWithAlpha() {
    if (!app.documents.length) {
        alert("沒有打開的文檔!");
        return;
    }
    var doc = app.activeDocument;
    
    // 解鎖背景圖層（如果存在），但保持其可見性狀態
    var bottomLayer = doc.layers[doc.layers.length - 1];
    var wasVisible = bottomLayer.visible;  // 記錄原始可見性狀態
    if (bottomLayer.isBackgroundLayer) {
        bottomLayer.isBackgroundLayer = false;
    }
    bottomLayer.visible = wasVisible;  // 恢復原始可見性狀態
    
    var originalName = doc.name.replace(/\.[^\.]+$/, '');
    
    // 檢查文檔是否已保存，如果沒有，使用桌面作為默認路徑
    var originalPath;
    try {
        originalPath = doc.path;
    } catch(e) {
        originalPath = Folder.desktop;
    }
    
    var duppedDocument = doc.duplicate();
    app.activeDocument = duppedDocument;
    
    // 確保總是有至少兩個圖層
    if (duppedDocument.layers.length === 1) {
        // 如果只有一個圖層，添加一個新的空白圖層
        var newLayer = duppedDocument.artLayers.add();
        newLayer.name = "合併圖層";
    }

    // 選擇所有圖層並合併
    selectAllLayers();
    duppedDocument.activeLayer.merge();

    // 確保合併後的圖層可見性為打開狀態
    duppedDocument.activeLayer.visible = true;

    function selectAllLayers() {
        // 選擇所有圖層（不包括背景）
        try {
            var desc = new ActionDescriptor();
            var ref = new ActionReference();
            ref.putEnumerated( charIDToTypeID('Lyr '), charIDToTypeID('Ordn'), charIDToTypeID('Trgt') );
            desc.putReference( charIDToTypeID('null'), ref );
            executeAction( stringIDToTypeID('selectAllLayers'), desc, DialogModes.NO );
        } catch(e) {}
        // 將背景圖層添加到選擇中（如果可能）
        try {
            activeDocument.backgroundLayer;
            var bgID = activeDocument.backgroundLayer.id;
            var ref = new ActionReference();
            var desc = new ActionDescriptor();
            ref.putIdentifier(charIDToTypeID('Lyr '), bgID);
            desc.putReference(charIDToTypeID('null'), ref);
            desc.putEnumerated( stringIDToTypeID('selectionModifier'), stringIDToTypeID('selectionModifierType'), stringIDToTypeID('addToSelection') );
            desc.putBoolean(charIDToTypeID('MkVs'), false);
            executeAction(charIDToTypeID('slct'), desc, DialogModes.NO);
        } catch(e) {}
    }
    
    function makeAlpha_from_Transparency() {
        try {
            var idSetd = charIDToTypeID("setd");
            var descSetd = new ActionDescriptor();
            var idNull = charIDToTypeID("null");
            var refFsel = new ActionReference();
            var idChnl = charIDToTypeID("Chnl");
            var idFsel = charIDToTypeID("fsel");
            refFsel.putProperty(idChnl, idFsel);
            descSetd.putReference(idNull, refFsel);
            var idTo = charIDToTypeID("T   ");
            var refChnl = new ActionReference();
            var idChnlEnum = charIDToTypeID("Chnl");
            var idTrsp = charIDToTypeID("Trsp");
            refChnl.putEnumerated(idChnlEnum, idChnlEnum, idTrsp);
            descSetd.putReference(idTo, refChnl);
            executeAction(idSetd, descSetd, DialogModes.NO);
            
            var idMk = charIDToTypeID("Mk  ");
            var descMk = new ActionDescriptor();
            var idNw = charIDToTypeID("Nw  ");
            var idChnlClass = charIDToTypeID("Chnl");
            descMk.putClass(idNw, idChnlClass);
            var idAt = charIDToTypeID("At  ");
            var refNewChnl = new ActionReference();
            var idChnlRef = charIDToTypeID("Chnl");
            var idNew = charIDToTypeID("New ");
            refNewChnl.putEnumerated(idChnlRef, idChnlRef, idNew);
            descMk.putReference(idAt, refNewChnl);
            var idUsng = charIDToTypeID("Usng");
            var idUsrM = charIDToTypeID("UsrM");
            var idRvlS = charIDToTypeID("RvlS");
            descMk.putEnumerated(idUsng, idUsrM, idRvlS);
            executeAction(idMk, descMk, DialogModes.NO);
            
            var alphaChannel = duppedDocument.channels[duppedDocument.channels.length - 1];
            duppedDocument.selection.store(alphaChannel, SelectionType.REPLACE);
        } catch (e) {
            alert("創建 alpha 通道時出錯: " + e);
        }
    }
    
    makeAlpha_from_Transparency();
    
    // 使用 Document 對象的 saveAs 方法保存為 TGA
    function saveTGA() {
        var saveOptions = new TargaSaveOptions();
        saveOptions.alphaChannels = true;
        saveOptions.resolution = TargaBitsPerPixels.THIRTYTWO;
        
        // 獲取上次保存的路徑
        var lastSavePath = getLastSavePath();
        var defaultPath;
        if (lastSavePath && Folder(lastSavePath).exists) {
            defaultPath = new File(lastSavePath + "/" + originalName + ".tga");
        } else {
            defaultPath = new File(originalPath + "/" + originalName + ".tga");
        }
        
        // 打開保存對話框
        var saveFile = defaultPath.saveDlg("保存為 TGA", "TGA:*.tga");
        
        if (saveFile) {
            duppedDocument.saveAs(saveFile, saveOptions, true, Extension.LOWERCASE);
            // 保存這次的路徑
            setLastSavePath(saveFile.parent.fsName);
        } else {
            alert("保存已取消");
        }
    }
    
    saveTGA();
    duppedDocument.close(SaveOptions.DONOTSAVECHANGES);
    app.activeDocument = doc;
}
// 執行腳本
saveAsTGAWithAlpha();