/*
    author: @37dmk and @fabiantheblind
    some repetitive changes you will want to do when typesetting big amounts of Text in inDesign.
*/


var save = confirm("Haben Sie Ihr Dokument schon gespeichert?");
if(!save){
    alert("Bitte speichern Sie Ihr Dokument bevor Sie fortfahren, damit Sie gegebenenfalls zur letzten Version zurückkehren können.");
    exit();
}

var myDialog;
with(myDialog = app.dialogs.add({name:"Mikrotypografie"})){
    with(dialogColumns.add()){
        with(dialogRows.add()){
            staticTexts.add({staticLabel:"Dieses Script ändert keine ausgeblendeten oder gesperrten Objekte."});
            staticTexts.add({staticLabel:" "});
            }
        with (dialogRows.add()) {
            with(dialogColumns.add()){
                with(borderPanels.add()){
                    staticTexts.add({staticLabel:"Sprachregion:"});
                    with(dialogColumns.add()){
                        var myRegion = dropdowns.add( { stringList: ["Schweiz", "Frankreich", "Deutschland", "England"], selectedIndex: 0 } );
                    }
                }
                with(borderPanels.add()){
                    staticTexts.add({staticLabel:"Änderungen:"});
                    with(dialogColumns.add()){
                        var strichlaengen = checkboxControls.add({staticLabel:"Strichlängen", checkedState:true});
                        var wortabstand = checkboxControls.add({staticLabel:"Doppelte Wortabstände löschen", checkedState:true});
                        var anfuehrung = checkboxControls.add({staticLabel:"Anführungszeichen", checkedState:true});
                        staticTexts.add({staticLabel:"_____________________________________________"});
                        staticTexts.add({staticLabel:" "});
                        var abstaende = enablingGroups.add({staticLabel:"Abstände", checkedState:true});
                        with(abstaende){
                            with(dialogColumns.add()){
                                var daten = checkboxControls.add({staticLabel:"Nummerisches Datum einbeziehen", checkedState:true});
                                }
                            }
                        staticTexts.add({staticLabel:"_____________________________________________"});
                        staticTexts.add({staticLabel:" "});
                        var zahlentrennung = enablingGroups.add({staticLabel:"Zahlentrennung", checkedState:true});
                        with(zahlentrennung){
                            with(dialogColumns.add()){
                                with(zahlentrennungButtons = radiobuttonGroups.add()){
                                var achtelgeviert = radiobuttonControls.add({staticLabel:"10 000            ", checkedState:true});
                                var apostroph = radiobuttonControls.add({staticLabel:"10’000"});
                                }
                                }
                            }
                        staticTexts.add({staticLabel:"_____________________________________________"});
                        staticTexts.add({staticLabel:" "});
                        var zeichensetzung = enablingGroups.add({staticLabel:"Zeichensetzung", checkedState:true});
                        with(zeichensetzung){
                            with(dialogColumns.add()){
                                var otf_hochgestellt = checkboxControls.add({staticLabel:"OTF hochgestellt", checkedState:true});
                                var otf_tiefgestellt = checkboxControls.add({staticLabel:"OTF tiefgestellt", checkedState:false});
                                var register_hochgestellt = checkboxControls.add({staticLabel:"'registered' hochgestellt", checkedState:false});
                                var copyright_hochgestellt = checkboxControls.add({staticLabel:"'copyright' hochgestellt", checkedState:true});
                                }
                            }
                        }
                    } // ende dialogColums
                } // ende border Panel
            } // ende Column
        } // ende Row
    }


myReturn = myDialog.show();
if (myReturn){
    var checkedCategories = new Array();
    if (anfuehrung.checkedState) {checkedCategories.push("Anführungszeichen")};
    if (strichlaengen.checkedState) {checkedCategories.push("Strichlängen")};
    if (abstaende.checkedState) {checkedCategories.push("Abstände")};
    if (zahlentrennung.checkedState) {checkedCategories.push("Zahlentrennung")};
    if (zeichensetzung.checkedState) {checkedCategories.push("Zeichensetzung")};
    if (wortabstand.checkedState) {checkedCategories.push("Doppelte Wortabstände")};
    var str = checkedCategories.join(", ");
    if (checkedCategories.length == 0) {
        alert("Bitte wählen Sie eine Änderung aus");
        }
    } else {
        myDialog.destroy()
        alert("Keine Änderungen vorgenommen");
        exit();
        }

var regionChosen = myRegion.stringList[myRegion.selectedIndex];

var doc = app.documents.item(0);
app.findGrepPreferences = NothingEnum.nothing;
app.changeGrepPreferences = NothingEnum.nothing;
app.findTextPreferences = NothingEnum.nothing;
app.changeTextPreferences = NothingEnum.nothing;
app.findChangeGrepOptions.includeFootnotes = true;
app.findChangeGrepOptions.includeHiddenLayers = false;
app.findChangeGrepOptions.includeLockedLayersForFind = false;
app.findChangeGrepOptions.includeLockedStoriesForFind = true;
app.findChangeGrepOptions.includeMasterPages = true;

var ztrenn_achtel_arr = [
{"findWhat":"(?<=\\d)'(?=\\d)","changeTo":"~<"},
{"findWhat":"(?<!\\d\\d)~<(?=\\d\\d\\d )","changeTo":".-.-.-.*-ç.-.-."},
{"findWhat":".-.-.-.*-ç.-.-.","changeTo":""},
]

var ztrenn_apos_arr = [
{"findWhat":"(?<=\\d)'(?=\\d)|(?<=\\d)~<(?=\\d)","changeTo":"’"},
{"findWhat":"(?<!\\d\\d)'(?=\\d\\d\\d )","changeTo":".-.-.-.*-ç.-.-."},
{"findWhat":".-.-.-.*-ç.-.-.","changeTo":""},
]

var anfuehr_ch_arr = [
{"findWhat":"\"(?![ ]|\\r|\\n|\\t|[.]|[,]|[;]|[:]|[!]|[?]|-|[)]|]|~S|~<|~|)","changeTo":"«~|"},
{"findWhat":"(?<![ ]|[(]|[[]|~S|~<|~||\\n|\\r)\"(?!~|)","changeTo":"~|»"},
{"findWhat":"«~|(?!.)","changeTo":"~|»"},
{"findWhat":"(?<![\\l\\u]|\\d)'(?! |\\r|\\n|\\t|\\.|\\,|\\;|\\:|\\!|\\?|-|\\(|\\[|~S|~<|~|)","changeTo":"‹~|"},
{"findWhat":"(?<![ ]|\\r|\\n|\\t|[.]|[,]|[;]|[:]|[!]|[?]|-|[(]|[[]|~S|~<|~|)'(?!\\d|\\l|\\u|~|)","changeTo":"~|›"},
{"findWhat":"(?<=[\\l\\u])'(?=\\l|\\u)","changeTo":"’"},
]

var anfuehr_fr_arr = [
{"findWhat":"\"(?![ ]|\\r|\\n|\\t|[.]|[,]|[;]|[:]|[!]|[?]|-|[)]|]|~S|~<|~|)","changeTo":"«~<"},
{"findWhat":"(?<![ ]|[(]|[[]|~S|~<|~||\\n|\\r)\"(?!~|)","changeTo":"~<»"},
{"findWhat":"«~|(?!.)","changeTo":"~<»"},
{"findWhat":"(?<![\\l\\u]|\\d)'(?! |\\r|\\n|\\t|\\.|\\,|\\;|\\:|\\!|\\?|-|\\(|\\[|~S|~<|~|)","changeTo":"‹~<"},
{"findWhat":"(?<![ ]|\\r|\\n|\\t|[.]|[,]|[;]|[:]|[!]|[?]|-|[(]|[[]|~S|~<|~|)'(?!\\d|\\l|\\u|~|)","changeTo":"~<›"},
{"findWhat":"(?<=[\\l\\u])'(?=\\l|\\u)","changeTo":"’"},
]

var anfuehr_de_arr = [
{"findWhat":"\"(?![ ]|\\r|\\n|\\t|[.]|[,]|[;]|[:]|[!]|[?]|-|[)]|]|~S|~<|~|)","changeTo":"»"},
{"findWhat":"(?<![ ]|[(]|[[]|~S|~<|~||\\n|\\r)\"(?!\\l|\\u)","changeTo":"«"},
{"findWhat":"»(?!.)","changeTo":"«"},
{"findWhat":"(?<![\\l\\u]|\\d)'(?! |\\r|\\n|\\t|\\.|\\,|\\;|\\:|\\!|\\?|-|\\(|\\[|~S|~<|~|)","changeTo":"›"},
{"findWhat":"(?<![ ]|\\r|\\n|\\t|[.]|[,]|[;]|[:]|[!]|[?]|-|[(]|[[]|~S|~<|~|)'(?!\\d|\\l|\\u|~|)","changeTo":"‹"},
{"findWhat":"(?<=[\\l\\u])'(?=\\l|\\u)","changeTo":"’"},
]

var anfuehr_en_arr = [
{"findWhat":"\"(?![ ]|\\r|\\n|\\t|[.]|[,]|[;]|[:]|[!]|[?]|-|[)]|]|~S|~<|~|)","changeTo":"“"},
{"findWhat":"(?<![ ]|[(]|[[]|~S|~<|~||\\n|\\r)\"(?!\\l|\\u)","changeTo":"”"},
{"findWhat":"“(?!.)","changeTo":"”"},
{"findWhat":"(?<![\\l\\u]|\\d)'(?! |\\r|\\n|\\t|\\.|\\,|\\;|\\:|\\!|\\?|-|\\(|\\[|~S|~<|~|)","changeTo":"‘"},
{"findWhat":"(?<![ ]|\\r|\\n|\\t|[.]|[,]|[;]|[:]|[!]|[?]|-|[(]|[[]|~S|~<|~|)'(?!\\d|\\l|\\u|~|)","changeTo":"’"},
{"findWhat":"(?<=[\\l\\u])'(?=\\l|\\u)","changeTo":"’"},
]

var strich_arr = [
{"findWhat":"(?<=\\d) - (?=\\d)|(?<=\\d[.]) - (?=\\d)|(?<=\\d[.])-(?=\\d)","changeTo":"~|~=~|"},
{"findWhat":"(?<=\\d)~=(?=\\d)","changeTo":"~|~=~|"},
{"findWhat":"(?<=\\d)-(?=\\d)","changeTo":"~|-~|"},
{"findWhat":" - ","changeTo":"~S~= "},
{"findWhat":" ~= ","changeTo":"~S~= "},
{"findWhat":"(?<=\\d[.])-","changeTo":"~="},
]

var zeichen_arr = [
{"findWhat":" x |(?<=\\d)x(?=\\d)|~<x~<","changeTo":"~%×~%"},
{"findWhat":"1/2","changeTo":"½"},
{"findWhat":"1/4","changeTo":"¼"},
{"findWhat":"3/4","changeTo":"¾"},
{"findWhat":"[.][.][.]","changeTo":"…"},
]

var abst_arr = [
{"findWhat":"(?<=\\d) %|(?<=\\d)%","changeTo":"~|%"},
{"findWhat":"(?<=\\d) (?=m|dm|km|cm|kg)","changeTo":"~%"},
{"findWhat":"(?<=\\d)°(?=C)|(?<=\\d) °(?=C)|(?<=\\d)° (?=C)|(?<=\\d) ° (?=C)","changeTo":"~S°~|"},
{"findWhat":"(?<=\\d)°(?=[ ])","changeTo":"~|°"},
{"findWhat":"[*](?=\\d)|[*] (?=\\d)","changeTo":"*~|"},
{"findWhat":"[†](?=\\d)|[†] (?=\\d)","changeTo":"†~|"},
{"findWhat":"(?<=\\d[.]) (?=Januar|Februar|März|April|Mai|Juni|Juli|August|September|Oktober|November|Dezember|Jahr|Liga|Platz|Rang|Preis|Mannschaft|Jh[.]|Jt[.]|Lebensjahr)","changeTo":"~%"},
{"findWhat":"(?<=Janua)r(?=\\d)","changeTo":"r~S"},
{"findWhat":"(?<=Februa)r(?=\\d)","changeTo":"r~S"},
{"findWhat":"(?<=Mär)z(?=\\d)","changeTo":"z~S"},
{"findWhat":"(?<=Apri)l(?=\\d)","changeTo":"l~S"},
{"findWhat":"(?<=Ma)i(?=\\d)","changeTo":"i~S"},
{"findWhat":"(?<=Jun)i(?=\\d)","changeTo":"i~S"},
{"findWhat":"(?<=Jul)i(?=\\d)","changeTo":"i~S"},
{"findWhat":"(?<=Augus)t(?=\\d)","changeTo":"t~S"},
{"findWhat":"(?<=Septembe)r(?=\\d)","changeTo":"r~S"},
{"findWhat":"(?<=Oktobe)r(?=\\d)","changeTo":"r~S"},
{"findWhat":"(?<=Novembe)r(?=\\d)","changeTo":"r~S"},
{"findWhat":"(?<=Dezembe)r(?=\\d)","changeTo":"r~S"},
{"findWhat":"(?<=Dr[.]) |(?<=med[.]) |(?<=Med[.]) |(?<=phil[.]) |(?<=lic[.]) |(?<=jur[.]) |(?<=iur[.]) |(?<=rer[.]) |(?<=pol[.]) |(?<=Prof[.]) |(?<=prof[.]) |(?<=St[.]) |(?<=Dipl[.]) |(?<=dipl[.]) |(?<=Nr[.]) |(?<=Fr[.]) |(?<=str[.]) |(?<=Ing[.]) |(?<=vet[.]) |(?<=Stv[.]) |(?<=Jh[.]) |(?<=Jt[.]) |(?<=bzw[.]) |(?<=ca[.]) ","changeTo":"~%"},
{"findWhat":"(?<=u[.]|U[.]) (?=a)|(?<=n[.]|N[.]|v[.]|V[.]) (?=Chr)|(?<=z[.]|Z[.]) (?=B)|(?<=d[.]|D[.]) (?=h)","changeTo":"~<"},
{"findWhat":"(?<=u[.]~<a[.]|U[.]~<a[.]) |(?<=Chr[.]|Chr[.]) |(?<=z[.]~<B[.]|Z[.]~<B[.]) |(?<=d[.]~<h[.]|D[.]~<h[.]) ","changeTo":"~%"},
{"findWhat":"(?<=\\d)[:](?=\\d)","changeTo":"~|:~|"},
{"findWhat":" […]","changeTo":"~S…"},
{"findWhat":"(?<=[\\l\\u]|\\d)[…]","changeTo":"~|…"},
{"findWhat":"(?<!~|)[/](?!~|)","changeTo":"~|/~|"},
{"findWhat":"(?<!~|)[\](?!~|)","changeTo":"~|\~|"},
{"findWhat":"[(](?!~|)","changeTo":"(~|"},
{"findWhat":"(?!~|)[)]","changeTo":"~|)"},
{"findWhat":"regex","changeTo":"newregex"},
]

var abst_ch_arr = [
{"findWhat":"(?<![ ]|~S|\\r|\\n|\\t|~<|~|)[!]","changeTo":"~|!"},
{"findWhat":"(?<![ ]|~S|\\r|\\n|\\t|~<|~|)[?]","changeTo":"~|?"},
{"findWhat":"(?<![ ]|~S|\\r|\\n|\\t|~<|~|)[:](?!\\d)","changeTo":"~|:"},
{"findWhat":"(?<![ ]|~S|\\r|\\n|\\t|~<|~|)[;]","changeTo":"~|;"},
]

var abst_fr_arr = [
{"findWhat":"(?<![ ]|~S|\\r|\\n|\\t|~<|~|)[!]","changeTo":"~<!"},
{"findWhat":"(?<![ ]|~S|\\r|\\n|\\t|~<|~|)[?]","changeTo":"~<?"},
{"findWhat":"(?<![ ]|~S|\\r|\\n|\\t|~<|~|)[:](?!\\d)","changeTo":"~<:"},
{"findWhat":"(?<![ ]|~S|\\r|\\n|\\t|~<|~|)[;]","changeTo":"~<;"},
]

var part_arr = [
{"findWhat":"regex","changeTo":"newregex"},
]

function singleChange(what, to, pos){
    app.findGrepPreferences.findWhat = what;
    if(to){
        app.changeGrepPreferences.changeTo = to;
    }
    if(pos){
        app.changeGrepPreferences.position = pos;
    }
    doc.changeGrep();
    app.findGrepPreferences = NothingEnum.nothing;
    app.changeGrepPreferences = NothingEnum.nothing;
}

function singleChangeText(what, to){
    app.findTextPreferences.findWhat = what;
    app.changeTextPreferences.changeTo = to;
    doc.changeText();
    app.findTextPreferences = NothingEnum.nothing;
    app.changeTextPreferences = NothingEnum.nothing;
}

function forChangeLoop(arr){
    for(var i=0; i < arr.length; i++){
        singleChange(arr[i].findWhat, arr[i].changeTo);
    }
}

if (achtelgeviert.checkedState) {
    forChangeLoop(ztrenn_achtel_arr);
    }

if (apostroph.checkedState) {
    forChangeLoop(ztrenn_apos_arr);
    }

if (anfuehrung.checkedState) {
    if (regionChosen == "Schweiz") {
        forChangeLoop(anfuehr_ch_arr);
        }
    if (regionChosen == "Frankreich") {
        forChangeLoop(anfuehr_fr_arr);
        }
    if (regionChosen == "Deutschland") {
        forChangeLoop(anfuehr_de_arr);
        }
    if (regionChosen == "England") {
        forChangeLoop(anfuehr_en_arr);
        }
    }

if (strichlaengen.checkedState) {
    forChangeLoop(strich_arr);
    }

if (zeichensetzung.checkedState) {
    forChangeLoop(zeichen_arr);
    if (regionChosen == "England"){
        if (otf_hochgestellt.checkedState){
            singleChange("(?<=\\d)st|(?<=\\d)nd|(?<=\\d)rd|(?<=\\d)th", 0, Position.otSuperscript);
            }
        else {
            singleChange("(?<=\\d)st|(?<=\\d)nd|(?<=\\d)rd|(?<=\\d)th", 0, Position.superscript);
            }
        }
    if (regionChosen == "Frankreich"){
        if (otf_hochgestellt.checkedState){
            singleChange("(?<=\\d)er|(?<=\\d)ème", 0, Position.otSuperscript);
            };
        else {
            singleChange("(?<=\\d)er|(?<=\\d)ème", 0, Position.superscript);
            };
        }
    if (otf_tiefgestellt.checkedState){
        singleChange("(?<=H)2(?=O)|(?<=H)12(?=O)|(?<=O)2|(?<=O)4|(?<=O)6|(?<=C)6", 0, Position.otSubscript);
        };
    else {
        singleChange("(?<=H)2(?=O)|(?<=H)12(?=O)|(?<=O)2|(?<=O)4|(?<=O)6|(?<=C)6", 0, Position.subscript);
        };
    if (otf_hochgestellt.checkedState){
        singleChange("(?<=m)2|(?<=m)3", 0, Position.otSuperscript);
        };
    else {
        singleChange("(?<=m)2|(?<=m)3", 0, Position.superscript);
        };
    if (register_hochgestellt.checkedState){
        singleChange("~r", "~|~r", Position.superscript);
        };
    if (copyright_hochgestellt.checkedState){
        singleChange("~2", "~|~2", Position.superscript);
        };
    if (regionChosen == "Schweiz") {
        singleChangeText("ß", "ss");
        };
    }

if (abstaende.checkedState) {
    forChangeLoop(abst_arr);
    if(daten.checkedState){
        singleChange("(?<=\\d)[.] (?=\\d)|(?<=\\d)[.](?=\\d\\d[.])|(?<=\\d)[.](?=\\d[.])|(?<=\\d)[.](?=\\d\\d\\d\\d)", ".~%");
        }
    if (regionChosen == "Frankreich") {
        forChangeLoop(abst_fr_arr);
    } else {
        forChangeLoop(abst_ch_arr);
    }
}

if (wortabstand.checkedState) {
    singleChange("    (?=[\\l\\u])|    (?=\\d)|    (?=[…])|    (?=[(])|    (?=\")|    (?=')|    (?=[*])|    (?=[†]|    (?=[|]))", "\\t");
    app.findChangeTextOptions.caseSensitive = false;
    app.findChangeTextOptions.wholeWord = false;
    for(i=0; i<31; i++){
        singleChangeText("  ", " ");
    }
    singleChange(" (?=\\t)|(?<=\\t) ", "");
    }

alert("Änderungen vorgenommen" + "\r" + str);

/* if you have any advice or questions about this feel free to contact me to lukas.schiltknecht@gmail.com */
