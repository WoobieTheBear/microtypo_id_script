function startDialog(){
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
                  with( borderPanels.add() ){
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
    return myDialog;
}
