/* Drop target */
var _target = document.getElementById('drop');
var _file = document.getElementById('file');
var _grid = document.getElementById('grid');

/* Spinner */
var spinner;

var _workstart = function() { spinner = new Spinner().spin(_target); }
var _workend = function() { spinner.stop(); }

/* Alerts */
var _badfile = function() {
  alertify.alert('This file does not appear to be a valid Excel file.', function (){});
};

var _pending = function() {
  alertify.alert('Please wait until the current file is processed.', function(){});
};

var _large = function(len, cb) {
  alertify.confirm("This file is " + len + " bytes and may take a few moments", cb);
};

var _failed = function(e) {
  console.log(e, e.stack);
  alertify.alert('We unfortunately dropped the ball here' , function(){});
};

/* Make the buttons for the sheets */

var make_buttons = function(sheetnames, cb) {

  // var buttons = document.getElementById('buttons');
  // buttons.innerHTML = "";

   sheetnames.forEach(function(s,idx) {
    var btn = document.createElement('button');
    btn.type = 'button';
    btn.name = 'btn' + idx;
    btn.text = s;
    var txt = document.createElement('h3'); txt.innerText = s; btn.appendChild(txt);
    btn.addEventListener('click', function() { cb(idx); }, false);

  //   buttons.appendChild(btn);
  //   buttons.appendChild(document.createElement('br'));

   });
};

var cdg = canvasDatagrid({
  parentNode: _grid
});


function loadFile(url,callback){
    JSZipUtils.getBinaryContent(url,callback);
}

function _resize() {
  _grid.style.height = (window.innerHeight - 200) + "px";
  _grid.style.width = (window.innerWidth - 200) + "px";
}
// function getNumber(){
//     jQuery.ajax({
//         url: "/num.php",
//         async: false
//     }).done(function( data ) {
//         res = data;
//     });
//     return res;
// }


function sendDoc(fio, docData){
    jQuery.ajax({
        url: "/recording.php",
        type: "POST",
        data: { fio: fio, data: docData },
        async: false
    }).done(function(data) {
        res = data;
    });
    return res;
}
window.addEventListener('resize', _resize);

var _onsheet = function(json, sheetnames, select_sheet_cb) {

  make_buttons(sheetnames, select_sheet_cb);

  /* show grid */
  _grid.style.display = "none";
  _resize();




  /* load data */
  cdg.data = json;
    sel = document.getElementById("city"); // Получаем наш список
    txt = sel.options[sel.selectedIndex].text;
    // noinspection JSAnnotator
    name = cdg.data[4][1] + " " + cdg.data[4][2] + " " + cdg.data[4][3];
    position = cdg.data[4][18] + " - " + cdg.data[4][19];
    //var secondName = cdg.data[4][1] + " " + cdg.data[4][2] + " " + cdg.data[4][3];
    passport = cdg.data[4][5] + " " + cdg.data[4][6] + " " + cdg.data[4][7];
    birthDate = cdg.data[4][4];
    address = cdg.data[4][8] + "; " + cdg.data[4][9] + "; " + cdg.data[4][10];
    gps = cdg.data[4][12];


    docData = "город - " + txt + "\n\r" +
        "паспортные данные - " + passport + "\n\r";
    //numberDocSecond
    numberDoc = sendDoc(name, docData);




     /* parse from the massive */
/*
    document.getElementById('secondCity').innerHTML = txt;
    document.getElementById('name').innerHTML = cdg.data[4][1] + " " + cdg.data[4][2] + " " + cdg.data[4][3];
    document.getElementById('position').innerHTML = cdg.data[4][18] + " - " + cdg.data[4][19];
    document.getElementById('secondName').innerHTML = cdg.data[4][1] + " " + cdg.data[4][2] + " " + cdg.data[4][3];
    document.getElementById('passport').innerHTML = cdg.data[4][5] + " " + cdg.data[4][6] + " " + cdg.data[4][7];
    document.getElementById('birthDate').innerHTML = cdg.data[4][4];
    document.getElementById('adress').innerHTML = cdg.data[4][8] + "; " + cdg.data[4][9] + "; " + cdg.data[4][10];
    document.getElementById('gps').innerHTML = cdg.data[4][12];


    console.log("text");
    console.log(cdg.data);
    */
    //console.log(txt);
  /* set up table headers */

  var L = 0;
  json.forEach(function(r) { if(L < r.length) L = r.length; });
  for(var i = 0; i < L; ++i) {

    cdg.schema[i].title = XLSX.utils.encode_col(i);
  }




    loadFile("/dogovor.docx",function(error,content){
        if (error) { throw error };
        var zip = new JSZip(content);
        var doc=new Docxtemplater().loadZip(zip)
        doc.setData({ // все наши поля
            city: txt,
            firstName: name,
            positionSecond: position + "-" + numberDoc,
            passportSecond: passport,
            birthSecond: birthDate,
            adressSecond: address,
            gpsSecond: gps
        });

        try {
            // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
            doc.render()
        }
        catch (error) {
            var e = {
                message: error.message,
                name: error.name,
                stack: error.stack,
                properties: error.properties,
            }
            console.log(JSON.stringify({error: e}));

            // The error thrown here contains additional information when logged with JSON.stringify (it contains a property object).

            throw error;
        }

        var out=doc.getZip().generate({
            type:"blob",
            mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        }) //Output the document using Data-URI
        saveAs(out,"contract.docx")
    })

};

/** Drop it like it's hot **/
DropSheet({
  file: _file,
  drop: _target,
  on: {
    workstart: _workstart,
    workend: _workend,
    sheet: _onsheet,
    foo: 'bar'
  },
  errors: {
    badfile: _badfile,
    pending: _pending,
    failed: _failed,
    large: _large,
    foo: 'bar'
  }
})
