/*
    author: @37dmk and @fabiantheblind
    some repetitive changes you will want to do if typesetting big amounts of Text.
*/


alert("Bitte speichern Sie Ihr Dokument bevor Sie fortfahren.");

var myDialog;
with(myDialog = app.dialogs.add({name:"Mikrotypografie"})){
    with(dialogColumns.add()){
        with(dialogRows.add()){
            staticTexts.add({staticLabel:"Dieses Script ändert keine ausgeblendeten oder gesperrten Objekte"});
            staticTexts.add({staticLabel:" "});
            }
        with (dialogRows.add()) {
            with(dialogColumns.add()){
                with(borderPanels.add()){
                    staticTexts.add({staticLabel:"Änderungen:"});
                    with(dialogColumns.add()){
                        var strichlaengen = checkboxControls.add({staticLabel:"Strichlängen", checkedState:true});
                        var wortabstand = checkboxControls.add({staticLabel:"Doppelte Wortabstände löschen", checkedState:true});
                        staticTexts.add({staticLabel:"_____________________________________________"});
                        staticTexts.add({staticLabel:" "});
                        var anfuehrung = enablingGroups.add({staticLabel:"Anführungszeichen", checkedState:true});
                        with(anfuehrung){
                            with(anfuehrungsButtons = radiobuttonGroups.add()){
                                var schweiz = radiobuttonControls.add({staticLabel:"schweizer/französische Guillemets", checkedState:true});
                                var deutsch = radiobuttonControls.add({staticLabel:"deutsche Guillemets"});
                                var englisch = radiobuttonControls.add({staticLabel:"englische Apostroph"});
                                }
                            }
                        staticTexts.add({staticLabel:"_____________________________________________"});
                        staticTexts.add({staticLabel:" "});
                        var abstaende = enablingGroups.add({staticLabel:"Abstände", checkedState:true});
                        with(abstaende){
                            with(dialogColumns.add()){
                                var daten = checkboxControls.add({staticLabel:"Nummerisches Datum einbeziehen", checkedState:true});
                                with(abstandButtons = radiobuttonGroups.add()){
                                var schweizabstand = radiobuttonControls.add({staticLabel:"schweizer Interpunktionsabstände", checkedState:true});
                                var franzabstand = radiobuttonControls.add({staticLabel:"französische Interpunktionsabstände"});
                                }
                                }
                            }
                        staticTexts.add({staticLabel:"_____________________________________________"});
                        staticTexts.add({staticLabel:" "});
                        var zahlentrennung = enablingGroups.add({staticLabel:"Zahlentrennung", checkedState:true});
                        with(zahlentrennung){
                            with(dialogColumns.add()){
                                with(zahlentrennungButtons = radiobuttonGroups.add()){
                                var achtelgeviert = radiobuttonControls.add({staticLabel:"10 000           ", checkedState:true});
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
                                var register_hochgestellt = checkboxControls.add({staticLabel:"'registered' hochgestellt", checkedState:true});
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
if (myReturn == true){
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
    }
    else {
        myDialog.destroy()
        alert("Keine Änderungen vorgenommen");
        exit();
        }


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
{"findWhat":"\"(?![ ]|\\r|\\n|\\t|[.]|[,]|[;]|[:]|[!]|[?]|-|[)]|]|~S|~|)","changeTo":"«~|"},
{"findWhat":"(?<![ ]|[(]|[[]|~S|~||\\n|\\r)\"(?!~|)","changeTo":"~|»"},
{"findWhat":"«~|(?!.)","changeTo":"~|»"},
{"findWhat":"(?<![\\l\\u]|\\d)'(?! |\\r|\\n|\\t|\\.|\\,|\\;|\\:|\\!|\\?|-|\\(|\\[|~S|~|)","changeTo":"‹~|"},
{"findWhat":"(?<![ ]|\\r|\\n|\\t|[.]|[,]|[;]|[:]|[!]|[?]|-|[(]|[[]|~S|~|)'(?!\\d|\\l|\\u|~|)","changeTo":"~|›"},
{"findWhat":"(?<=[\\l\\u])'(?=\\l|\\u)","changeTo":"’"},
]

var anfuehr_de_arr = [
{"findWhat":"\"(?![ ]|\\r|\\n|\\t|[.]|[,]|[;]|[:]|[!]|[?]|-|[)]|]|~S|~|)","changeTo":"»"},
{"findWhat":"(?<![ ]|[(]|[[]|~S|~||\\n|\\r)\"(?!\\l|\\u)","changeTo":"«"},
{"findWhat":"»(?!.)","changeTo":"«"},
{"findWhat":"(?<![\\l\\u]|\\d)'(?! |\\r|\\n|\\t|\\.|\\,|\\;|\\:|\\!|\\?|-|\\(|\\[|~S|~|)","changeTo":"›"},
{"findWhat":"(?<![ ]|\\r|\\n|\\t|[.]|[,]|[;]|[:]|[!]|[?]|-|[(]|[[]|~S|~|)'(?!\\d|\\l|\\u|~|)","changeTo":"‹"},
{"findWhat":"(?<=[\\l\\u])'(?=\\l|\\u)","changeTo":"’"},
]

var anfuehr_en_arr = [
{"findWhat":"\"(?![ ]|\\r|\\n|\\t|[.]|[,]|[;]|[:]|[!]|[?]|-|[)]|]|~S|~|)","changeTo":"“"},
{"findWhat":"(?<![ ]|[(]|[[]|~S|~||\\n|\\r)\"(?!\\l|\\u)","changeTo":"”"},
{"findWhat":"“(?!.)","changeTo":"”"},
{"findWhat":"(?<![\\l\\u]|\\d)'(?! |\\r|\\n|\\t|\\.|\\,|\\;|\\:|\\!|\\?|-|\\(|\\[|~S|~|)","changeTo":"‘"},
{"findWhat":"(?<![ ]|\\r|\\n|\\t|[.]|[,]|[;]|[:]|[!]|[?]|-|[(]|[[]|~S|~|)'(?!\\d|\\l|\\u|~|)","changeTo":"’"},
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

if (achtelgeviert.checkedState) {
    for(var i = 0;i < ztrenn_achtel_arr.length;i++){
        app.findGrepPreferences.findWhat = ztrenn_achtel_arr[i].findWhat;
        app.changeGrepPreferences.changeTo = ztrenn_achtel_arr[i].changeTo;
        doc.changeGrep();
        app.findGrepPreferences = NothingEnum.nothing;
        app.changeGrepPreferences = NothingEnum.nothing;
        };
    }

if (apostroph.checkedState) {
    for(var i = 0;i < ztrenn_apos_arr.length;i++){
        app.findGrepPreferences.findWhat = ztrenn_apos_arr[i].findWhat;
        app.changeGrepPreferences.changeTo = ztrenn_apos_arr[i].changeTo;
        doc.changeGrep();
        app.findGrepPreferences = NothingEnum.nothing;
        app.changeGrepPreferences = NothingEnum.nothing;
        };
    }

if (anfuehrung.checkedState) {
    if (schweiz.checkedState) {
        for(var i = 0;i < anfuehr_ch_arr.length;i++){
            app.findGrepPreferences.findWhat = anfuehr_ch_arr[i].findWhat;
            app.changeGrepPreferences.changeTo = anfuehr_ch_arr[i].changeTo;
            doc.changeGrep();
            app.findGrepPreferences = NothingEnum.nothing;
            app.changeGrepPreferences = NothingEnum.nothing;
            };
        }
    if (deutsch.checkedState) {
        for(var i = 0;i < anfuehr_de_arr.length;i++){
            app.findGrepPreferences.findWhat = anfuehr_de_arr[i].findWhat;
            app.changeGrepPreferences.changeTo = anfuehr_de_arr[i].changeTo;
            doc.changeGrep();
            app.findGrepPreferences = NothingEnum.nothing;
            app.changeGrepPreferences = NothingEnum.nothing;
            };
        }
    if (englisch.checkedState) {
        for(var i = 0;i < anfuehr_en_arr.length;i++){
            app.findGrepPreferences.findWhat = anfuehr_en_arr[i].findWhat;
            app.changeGrepPreferences.changeTo = anfuehr_en_arr[i].changeTo;
            doc.changeGrep();
            app.findGrepPreferences = NothingEnum.nothing;
            app.changeGrepPreferences = NothingEnum.nothing;
            };
        }
    }

if (strichlaengen.checkedState) {
    for(var i = 0;i < strich_arr.length;i++){
        app.findGrepPreferences.findWhat = strich_arr[i].findWhat;
        app.changeGrepPreferences.changeTo = strich_arr[i].changeTo;
        doc.changeGrep();
        app.findGrepPreferences = NothingEnum.nothing;
        app.changeGrepPreferences = NothingEnum.nothing;
        };    
    }

if (zeichensetzung.checkedState) {
    for(var i = 0;i < zeichen_arr.length;i++){
        app.findGrepPreferences.findWhat = zeichen_arr[i].findWhat;
        app.changeGrepPreferences.changeTo = zeichen_arr[i].changeTo;
        doc.changeGrep();
        app.findGrepPreferences = NothingEnum.nothing;
        app.changeGrepPreferences = NothingEnum.nothing;
        };
    if (englisch.checkedState){
        if (otf_hochgestellt.checkedState){
            app.findGrepPreferences.findWhat = "(?<=\\d)st|(?<=\\d)nd|(?<=\\d)rd|(?<=\\d)th";
            app.changeGrepPreferences.position = Position.otSuperscript;
            doc.changeGrep();
            app.findGrepPreferences = NothingEnum.nothing;
            app.changeGrepPreferences = NothingEnum.nothing;
            }
        else {
            app.findGrepPreferences.findWhat = "(?<=\\d)st|(?<=\\d)nd|(?<=\\d)rd|(?<=\\d)th";
            app.changeGrepPreferences.position = Position.superscript;
            doc.changeGrep();
            app.findGrepPreferences = NothingEnum.nothing;
            app.changeGrepPreferences = NothingEnum.nothing;
            }
        }
    if (schweiz.checkedState) {
        app.findTextPreferences.findWhat = "ß";
        app.changeTextPreferences.changeTo = "ss";
        doc.changeText();
        app.findTextPreferences = NothingEnum.nothing;
        app.changeTextPreferences = NothingEnum.nothing;
        };
    if (franzabstand.checkedState){
        if (otf_hochgestellt.checkedState){
            app.findGrepPreferences.findWhat = "(?<=\\d)er|(?<=\\d)ème";
            app.changeGrepPreferences.position = Position.otSuperscript;
            doc.changeGrep();
            app.findGrepPreferences = NothingEnum.nothing;
            app.changeGrepPreferences = NothingEnum.nothing;
            };
        else {
            app.findGrepPreferences.findWhat = "(?<=\\d)er|(?<=\\d)ème";
            app.changeGrepPreferences.position = Position.superscript;
            doc.changeGrep();
            app.findGrepPreferences = NothingEnum.nothing;
            app.changeGrepPreferences = NothingEnum.nothing;
            };
        }
    if (otf_tiefgestellt.checkedState){
        app.findGrepPreferences.findWhat = "(?<=H)2(?=O)|(?<=H)12(?=O)|(?<=O)2|(?<=O)4|(?<=O)6|(?<=C)6";
        app.changeGrepPreferences.position = Position.otSubscript;
        doc.changeGrep();
        app.findGrepPreferences = NothingEnum.nothing;
        app.changeGrepPreferences = NothingEnum.nothing;
        };
    else {
        app.findGrepPreferences.findWhat = "(?<=H)2(?=O)|(?<=H)12(?=O)|(?<=O)2|(?<=O)4|(?<=O)6|(?<=C)6";
        app.changeGrepPreferences.position = Position.subscript;
        doc.changeGrep();
        app.findGrepPreferences = NothingEnum.nothing;
        app.changeGrepPreferences = NothingEnum.nothing;
        };
    if (otf_hochgestellt.checkedState){
        app.findGrepPreferences.findWhat = "(?<=m)2|(?<=m)3";
        app.changeGrepPreferences.position = Position.otSuperscript;
        doc.changeGrep();
        app.findGrepPreferences = NothingEnum.nothing;
        app.changeGrepPreferences = NothingEnum.nothing;
        };
    else {
        app.findGrepPreferences.findWhat = "(?<=m)2|(?<=m)3";
        app.changeGrepPreferences.position = Position.superscript;
        doc.changeGrep();
        app.findGrepPreferences = NothingEnum.nothing;
        app.changeGrepPreferences = NothingEnum.nothing;
        };
    if (register_hochgestellt.checkedState){
        app.findGrepPreferences.findWhat = "~r";
        app.changeGrepPreferences.changeTo = "~|~r";
        app.changeGrepPreferences.position = Position.superscript;
        doc.changeGrep();
        app.findGrepPreferences = NothingEnum.nothing;
        app.changeGrepPreferences = NothingEnum.nothing;
        };
    if (copyright_hochgestellt.checkedState){
        app.findGrepPreferences.findWhat = "~2";
        app.changeGrepPreferences.changeTo = "~|~2";
        app.changeGrepPreferences.position = Position.superscript;
        doc.changeGrep();
        app.findGrepPreferences = NothingEnum.nothing;
        app.changeGrepPreferences = NothingEnum.nothing;
        };
    }

if (abstaende.checkedState) {
    for(var i = 0;i < abst_arr.length;i++){
        app.findGrepPreferences.findWhat = abst_arr[i].findWhat;
        app.changeGrepPreferences.changeTo = abst_arr[i].changeTo;
        doc.changeGrep();
        app.findGrepPreferences = NothingEnum.nothing;
        app.changeGrepPreferences = NothingEnum.nothing;
        };
    if(daten.checkedState){
        app.findGrepPreferences.findWhat = "(?<=\\d)[.] (?=\\d)|(?<=\\d)[.](?=\\d\\d[.])|(?<=\\d)[.](?=\\d[.])|(?<=\\d)[.](?=\\d\\d\\d\\d)";
        app.changeGrepPreferences.changeTo = ".~%";
        doc.changeGrep();
        app.findGrepPreferences = NothingEnum.nothing;
        app.changeGrepPreferences = NothingEnum.nothing;
        }
    if (schweizabstand.checkedState) {
        for(var i = 0;i < abst_ch_arr.length;i++){
            app.findGrepPreferences.findWhat = abst_ch_arr[i].findWhat;
            app.changeGrepPreferences.changeTo = abst_ch_arr[i].changeTo;
            doc.changeGrep();
            app.findGrepPreferences = NothingEnum.nothing;
            app.changeGrepPreferences = NothingEnum.nothing;
            };
        }
    if (schweizabstand.checkedState) {
        for(var i = 0;i < abst_fr_arr.length;i++){
            app.findGrepPreferences.findWhat = abst_fr_arr[i].findWhat;
            app.changeGrepPreferences.changeTo = abst_fr_arr[i].changeTo;
            doc.changeGrep();
            app.findGrepPreferences = NothingEnum.nothing;
            app.changeGrepPreferences = NothingEnum.nothing;
            };
        }
    }

if (wortabstand.checkedState) {
    app.findGrepPreferences.findWhat = "    (?=[\\l\\u])|    (?=\\d)|    (?=[…])|    (?=[(])|    (?=\")|    (?=')|    (?=[*])|    (?=[†]|    (?=[|]))";
    app.changeGrepPreferences.changeTo = "\\t";
    doc.changeGrep();
    app.findGrepPreferences = NothingEnum.nothing;
    app.changeGrepPreferences = NothingEnum.nothing;
    app.findChangeTextOptions.caseSensitive = false;
    app.findChangeTextOptions.wholeWord = false;
    app.findTextPreferences.findWhat = "  ";
    app.changeTextPreferences.changeTo = " ";
    var i=0
    do
    {
    doc.changeText();
    i++;
    }
    while (i<=30);
    app.findTextPreferences = NothingEnum.nothing;
    app.changeTextPreferences = NothingEnum.nothing;
    app.findGrepPreferences.findWhat = " (?=\\t)|(?<=\\t) ";
    app.changeGrepPreferences.changeTo = "";
    doc.changeGrep();
    app.findGrepPreferences = NothingEnum.nothing;
    app.changeGrepPreferences = NothingEnum.nothing;
    }

alert("Änderungen vorgenommen" + "\r" + str);

/* if you have any advice or questions about this feel free to cantact me to lukas.schiltknecht@gmx.ch */
