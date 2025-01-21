//Toggles the current displayed Div
function toggleDiv(divName, activeDiv) {
    try { 
        document.getElementById(activeDiv).style.display = "none";
        document.getElementById(divName).style.display = "inline";
    } catch (error) {
        console.error("Error in toggleDiv:", error);
    }
}

//Filters the table based on the filter field and hides other rows
function tableFilter(elementId, tableName) {
    document.getElementById(elementId).addEventListener('change', function() {
        var filter = this.value;
        var rows = document.querySelectorAll(`#${tableName} tbody tr`);
      
        rows.forEach(function(row) {
          var type = row.cells[0].textContent.toLowerCase();
          if (filter === 'all' || type === filter) {
            row.style.display = '';
          } else {
            row.style.display = 'none';
          }
        });
      });
}

//Rotates the sync icon on hover
function transformSync() {
    document.getElementById("syncIcon").style.transform = 'rotate(360deg)';
}

//Resets the animation on the sync icon
function resetSync() {
    document.getElementById("syncIcon").style.transform = 'rotate(0deg)';
}

//Shows family tree
function showFamilyTree(contact) {

    document.getElementById(`${contact}FamilyTreeGroup`).style.display = "block";
    document.getElementById(`popupBackground`).style.display = "block";
}

//Hides family tree
function hideFamilyTree(contact) {

    document.getElementById(`${contact}FamilyTreeGroup`).style.display = "none";
    document.getElementById(`popupBackground`).style.display = "none";
}

//Hides all divs besides overview to begin with
function setInitialFocus() {
    const elementIds = [
        {"id": "loading"},
        {"id": "overview"},
        {"id": "objectives"},
        {"id": "risk"},
        {"id": "dna"},
        {"id": "income"},
        {"id": "asset"},
        {"id": "prof"},
        {"id": "prof2"},
        {"id": "fnf"},
        {"id": "actions"},
        {"id": "syncing"},
        {"id": "tax"},
        {"id": "pcFamilyTreeGroup"},
        {"id": "scFamilyTreeGroup"},
        {"id": "popupBackground"},
        {"id": "trusts"},
        {"id": "companies"}
    ]

    elementIds.forEach(function(element){
        if (element.id == "loading") {
            document.getElementById(element.id).style.display = "inline";
        } else {
            document.getElementById(element.id).style.display = "none";
        }
    })

    //document.getElementById("syncButton").addEventListener('mouseover', transformSync);
    //document.getElementById("syncButton").addEventListener('mouseout', resetSync);
    firstLoad("loading", "loadText", "overview");
}

//Sets the window focus on a tab housed within the account
function tabFocusLocal(fieldName, elementId) {
    if (document.getElementById(elementId).textContent !== "") {
        console.log(document.getElementById(elementId).textContent)
        return;
    } else {
        control = window.parent.Xrm.Page.getControl(fieldName);
        section = control.getParent();
        tab = section.getParent().getName();


        var anchor = document.createElement("a");
        switch (tab) {
            case "Cash flow":
                anchor.textContent = "£0.00";
            break;

            default:
                anchor.textContent = "Set info...";
            break;
        }
        anchor.href = tab;
        anchor.id = tab;
        anchor.onclick = function(event) {
            event.preventDefault();
            jumpToLocalTab(anchor.id);
        }
        document.getElementById(elementId).appendChild(anchor);
    }
}

//Creates an iframe over the dashboard with a view into the related table
function tabFocusOrg(elementId, entityName, recordId, tabName) {

    if (document.getElementById(elementId).textContent !== "") {
        return;
    } else {
        var anchor = document.createElement("a");
        anchor.href = recordId;
        switch (tabName) {
            case "Risk Questionnaire":
                anchor.textContent = "Questionnaire not filled out";
            break;

            default:
                anchor.textContent = "Set info...";
            break;
        }
        anchor.id = recordId;
        anchor.onclick = function(event) {
            event.preventDefault();
            openFormWithTab(entityName, recordId, tabName);
        }  
        document.getElementById(elementId).appendChild(anchor);     
    }
}

//Creates table rows for tables
function createTableOfEntityLinkedToEntity(tableId, fieldName, recordId, entityName, rows, formId) {
    var query = `?$filter=${fieldName} eq ${encodeURIComponent("'" + recordId + "'")} and statecode eq 0`;
    console.log(query)
    window.parent.Xrm.WebApi.retrieveMultipleRecords(entityName, query).then(
        function (response) {
            var entities = response.entities;
            // Get a reference to the table
            var table = document.getElementById(tableId);
            console.log(table);
            // Clear existing rows (optional)
             while (table.rows.length >= 1) {
                 table.deleteRow(0);
            }

            if (entityName == "ax_income") {
                sortedData = entities.sort((a, b) => {
                    if (a["ax_category@OData.Community.Display.V1.FormattedValue"] < b["ax_category@OData.Community.Display.V1.FormattedValue"]) {
                      return -1;
                    }
                    if (a["ax_category@OData.Community.Display.V1.FormattedValue"] > b["ax_category@OData.Community.Display.V1.FormattedValue"]) {
                      return 1;
                    }
                    // If ax_category is the same, sort by ax_contact_value
                    if (a["_ax_contact_value@OData.Community.Display.V1.FormattedValue"] < b["_ax_contact_value@OData.Community.Display.V1.FormattedValue"]) {
                      return -1;
                    }
                    if (a["_ax_contact_value@OData.Community.Display.V1.FormattedValue"] > b["_ax_contact_value@OData.Community.Display.V1.FormattedValue"]) {
                      return 1;
                    }
                    return 0;
                  });
            } else if(entityName == "ax_asset"){
                sortedData = entities.sort((a, b) => {
                    if (a["ax_category@OData.Community.Display.V1.FormattedValue"] < b["ax_category@OData.Community.Display.V1.FormattedValue"]) {
                      return -1;
                    }
                    if (a["ax_category@OData.Community.Display.V1.FormattedValue"] > b["ax_category@OData.Community.Display.V1.FormattedValue"]) {
                      return 1;
                    }
                    // If ax_category is the same, sort by ax_contact_value
                    if (a["ax_assetowner"] < b["ax_assetowner"]) {
                      return -1;
                    }
                    if (a["ax_assetowner"] > b["ax_assetowner"]) {
                      return 1;
                    }
                    return 0;
                  });
            } else {
                sortedData = entities;
            }
            


            // Iterate over the entities and create table rows
            sortedData.forEach(function(income, index) {
                var newRow = table.insertRow();

                switch (entityName){
                    case "ax_income":
                        var rowId = income["ax_incomeid"];
                        break;
                    case "ax_asset":
                        var rowId = income["ax_assetid"];
                        break;
                    case "ax_trust":
                        var rowId = income["ax_trustid"];
                        break;
                    case "ax_company":
                        var rowId = income["ax_companyid"];
                        break;
                    case "ax_keyfact":
                        var rowId = income["ax_keyfactid"];
                        break;
                    case "ax_feedback":
                        var rowId = income["ax_feedbackid"];
                        break;
                    case "task":
                        var rowId = income["activityid"];
                    case "connection":
                        var rowId = income["_record2id_value"]
                        break;
                }

                switch (entityName) {
                    case "connection":
                        newRow.addEventListener('click', function(event) {
                            event.preventDefault();
                            openFormWithTab("contact", rowId, "Details");
                        })
                        break;
                    default:
                        newRow.addEventListener('click', function(event) {
                            event.preventDefault();
                            openFormWithTab(entityName, rowId, getTabNameOfField(entityName, formId, income[rows[1].field]));
                        })
                        break;
                }

                rows.forEach(function(row, cellIndex) {
                    var newCell = newRow.insertCell(cellIndex);
                    newCell.style.width = row.width;
                    switch (entityName) {
                      case "ax_income":
                        switch (row.field) {
                          case "ax_totalremuneration@OData.Community.Display.V1.FormattedValue":
                            if (income[rows[0].field] == "Employment") {
                              newCell.textContent = income[row.field] || "-";
                            } else {
                              newCell.textContent = income["ax_annual_amount@OData.Community.Display.V1.FormattedValue"] || "-";
                            }
                            break;
                          case "ax_companyname":
                            switch (income[rows[0].field]) {
                              case "Employment":
                                newCell.textContent = income[row.field] || "-";
                                break;
                              case "Asset based":
                                newCell.textContent = income["_ax_asset_value@OData.Community.Display.V1.FormattedValue"] || "-";
                                break;
                              case "Other":
                                newCell.textContent = income["ax_source"] || "-";
                                break;
                              default:
                                newCell.textContent = "-";
                            }
                            break;
                          default:
                            newCell.textContent = income[row.field] || "-";
                        }
                        break;
                      case "ax_asset":
                        switch (row.field) {
                            case "ax_category@OData.Community.Display.V1.FormattedValue":
                                var str = income["ax_category@OData.Community.Display.V1.FormattedValue"];
                                newCell.textContent = str.substring(str.indexOf(' ') + 1) || "-";
                                break;
                            default:
                                newCell.textContent = income[row.field] || "-";
                        }
                      break;
                      case "connection":
                        getContact(rowId).then(function(result) {
                            switch (row.field) {
                                case "age":
                                    newCell.textContent = result["ax_currentage"] || "-";
                                    break;
                                case "status":
                                    newCell.textContent = result["statuscode@OData.Community.Display.V1.FormattedValue"] || "-";
                                    break;
                                default:
                                newCell.textContent = income[row.field] || "-";
                    }})
                        
                      break;
                      default:
                        newCell.textContent = income[row.field] || "-";
                    }
                  });
                  
            });
        },
        function (error) {
            console.error("Error retrieving entities: " + error.message);
        }
    );
}

//sub function of tabFocusOrg
function openFormWithTab(entityName, recordId, tabname) {
    document.getElementById("loading").style.display = "block";
    var pageInput = {
        pageType: "entityrecord",
        entityName: entityName,
        entityId: recordId,
        formId: null,
        tabName: tabname
    };

    var navigationOptions = {
        target: 2,
    }
    console.log(tabname);
    var tabName;
    switch (entityName) {
        case "ax_asset":
            tabName = "asset";
            break;
        case "ax_trust":
            tabName = "trusts";
            break;
        case "ax_company":
            tabName = "companies";
            break;
        case "ax_income":
            tabName = "income";
            break;
        case "task":
            tabName = "actions";
            break;
        case "ax_feedback":
            tabName = "fnf";
            break;
        case "ax_keyfact":
            tabName = "fnf";
            break;
        case "contact":
            switch (tabname) {
                case "Tax Planning":
                    tabName = "tax";
                    break;
                default:
                    tabName = "overview"
                    break;
            }
            break;
        default:
            tabName = "overview";
            break;
    }

    window.parent.Xrm.Navigation.navigateTo(pageInput, navigationOptions).then(
        function (success) {
                    console.log("Form opened successfully.");
                    document.getElementById("loading").style.display = "none";
                    firstLoad("syncing", 'loadText', tabName);
                },
                function (error) {
                    console.error("Error opening form: " + error.message);
                }
            );
}

//sub function of tabFocusLocal
function jumpToLocalTab(tabName) {
    var formContext = window.parent.Xrm.Page;
    var tab = formContext.ui.tabs.get(tabName);
    
    if (tab) {
        tab.setFocus();
    } else {
        console.error("Tab not found: " + tabName);
    }
}

//Gets the value of any basic field
function getFieldValue(elementId, fieldName) {
    if (window.parent.Xrm.Page.getAttribute(fieldName).getValue()) {
        document.getElementById(elementId).textContent =  window.parent.Xrm.Page.getAttribute(fieldName).getValue();
        if (elementId == "background") {
            console.log(window.parent.Xrm.Page.getAttribute(fieldName).getValue());
        }
    } else {
        tabFocusLocal(fieldName, elementId)
    }
}

//Hard coded functionality to format the number as stars
function getFieldValueRating(elementId, fieldName) {
    if (window.parent.Xrm.Page.getAttribute(fieldName).getValue()) {
        switch(window.parent.Xrm.Page.getAttribute(fieldName).getValue()) {
            case 1:
                document.getElementById(elementId).textContent = "⭐";
                break;
            case 2:
                document.getElementById(elementId).textContent = "⭐⭐";
                break;
            case 3:
                document.getElementById(elementId).textContent = "⭐⭐⭐";
                break;
            default:
                document.getElementById(elementId).textContent = "null";
        }
    } else {
        tabFocusLocal(fieldName, elementId)
    }
}

//Gets the text value of a lookup field
function getLookupFieldName(elementId, fieldName) {
    if (window.parent.Xrm.Page.getAttribute(fieldName).getValue()) {
        document.getElementById(elementId).textContent =  window.parent.Xrm.Page.getAttribute(fieldName).getValue()[0].name;
    } else {
        tabFocusLocal(fieldName, elementId)
    }
}

//Gets the text value of an optionset option
function getOptionsetFieldName(elementId, fieldName) {
    if (window.parent.Xrm.Page.getAttribute(fieldName).getSelectedOption()) {
        document.getElementById(elementId).textContent =  window.parent.Xrm.Page.getAttribute(fieldName).getSelectedOption().text
    } else {
        tabFocusLocal(fieldName, elementId)
    }
}

//Formats currency in GBP
function formatCurrency(value, locale = 'en-GB', currency = 'GBP') {
    return new Intl.NumberFormat(locale, { style: 'currency', currency: currency }).format(value);
}

//Gets formatted value of currency fields
function getFormattedValueFieldName(elementId, fieldName) {
    if (window.parent.Xrm.Page.getAttribute(fieldName).getValue()) {
        var amount = window.parent.Xrm.Page.getAttribute(fieldName).getValue()
        var formattedAmount = formatCurrency(amount);
        document.getElementById(elementId).textContent =  formattedAmount;
    } else {
        tabFocusLocal(fieldName, elementId)
    }
}

//Hard coded for length of time for service
function getYearsAndDays(elementId, fieldName1, fieldName2) {
    if (window.parent.Xrm.Page.getAttribute(fieldName2).getValue()) {
        document.getElementById(elementId).textContent =  `${window.parent.Xrm.Page.getAttribute(fieldName1).getValue()} years and ${window.parent.Xrm.Page.getAttribute(fieldName2).getValue()} days`
    } else {
        tabFocusLocal(fieldName1, elementId)
    }
}

//Gets the tab name in which a field resides
function getTabNameOfField(entityLogicalName, formId, fieldName) {
    return new Promise((resolve, reject) => {
        var query = `/api/data/v9.1/systemforms(${formId})?$select=formxml`;

        window.parent.Xrm.WebApi.retrieveRecord("systemform", formId, "?$select=formxml").then(
            function (response) {
                var parser = new DOMParser();
                var xmlDoc = parser.parseFromString(response.formxml, "text/xml");
                var fieldNode = xmlDoc.querySelector(`control[datafieldname='${fieldName}']`);
                
                if (fieldNode) {
                    var tabNode = fieldNode.closest("tab");
                    if (tabNode) {
                        var tabName = tabNode.getAttribute("name");
                        resolve(tabName);
                    } else {
                        console.error("Tab not found for the field.");
                        reject("Tab not found for the field.");
                    }
                } else {
                    console.error("Field not found in the form.");
                    reject("Field not found in the form.");
                }
            },
            function (error) {
                console.error("Error retrieving form XML: " + error.message);
                reject(error.message);
            }
        );
    });
}

//Sets the value of a field
function apiField(elementId, recordName, recordId, recordType, fieldName) {
    if (recordName) {
        document.getElementById(elementId).textContent = recordName;
    } else {
        tab = getTabNameOfField(recordType, "c9563986-77e5-eb11-bacb-0022481a7d07", fieldName).then(
            function(tabName) {
                switch (tabName) {
                    case "Risk Questionnaire":
                        document.getElementById(elementId).style.display = "none";
                        document.getElementById(elementId).parentElement.style.display = "none";
                        document.getElementById("noQuest").style.display = "block";
                        break;
                    default:
                        tabFocusOrg(elementId, recordType, recordId, tabName);
                        break;
                }
            },
            function(error) {
                console.error(error);
            }
        );
    }
}

//Hard coded for all contact fields both primary and secondary determined via a prefix of sc or pc
function setContactFields(fieldName, prefix) {
    getContact(window.parent.Xrm.Page.getAttribute(fieldName).getValue()[0].id).then(function(result) {
        var dateOfBirth = result["ax_dateofbirth"];
        var dateOfBirthClean = new Date(dateOfBirth);
        var formattedDateOfBirth = dateOfBirthClean.toLocaleString('en-GB', {
            day: '2-digit',
            month: '2-digit',
            year: 'numeric'
        })

        var dateOfWill = result["ax_dateofwill"];
        var dateOfWillClean = new Date(dateOfWill);
        var formattedDateOfWill = dateOfWillClean.toLocaleString('en-GB', {
            day: '2-digit',
            month: '2-digit',
            year: 'numeric'
        })

        var dateFinaCompleted = result["ax_datefinametricawascompleted"];
        var dateFinaCompletedClean = new Date(dateFinaCompleted);
        var formattedDataFinaCompleted = dateFinaCompletedClean.toLocaleString('en-GB', {
            day: '2-digit',
            month: '2-digit',
            year: 'numeric' 
        })

        apiField(`${prefix}KnownAs`, result["ax_knownas"], result["contactid"], "contact", "ax_knownas")
        apiField(`${prefix}Dob`, formattedDateOfBirth.concat(` - ${result["ax_currentage"]}`), result["contactid"], "contact", "ax_dateofbirth")
        apiField(`${prefix}Employment`, result["ax_employmentstatus@OData.Community.Display.V1.FormattedValue"], result["contactid"], "contact", "ax_employmentstatus")
        if (result["ax_residency@OData.Community.Display.V1.FormattedValue"] == "Other") {
            apiField(`${prefix}Residency`, result["ax_residentof"] , result["contactid"], "contact", "ax_residency")
        } else {
            apiField(`${prefix}Residency`, result["ax_residency@OData.Community.Display.V1.FormattedValue"] , result["contactid"], "contact", "ax_residency")
        }

        if (result["ax_domicile@OData.Community.Display.V1.FormattedValue"] == "Other") {
            apiField(`${prefix}Domicile`,result["ax_domicileof"] , result["contactid"], "contact", "ax_domicile")
        } else {
            apiField(`${prefix}Domicile`, result["ax_domicile@OData.Community.Display.V1.FormattedValue"], result["contactid"], "contact", "ax_domicile")

        }
        
        apiField(`${prefix}Nationality`, result["ax_nationality"], result["contactid"], "contact","ax_nationality")
        apiField(`${prefix}Name`, result["fullname"], result["contactid"], "contact", "fullname")
        apiField(`${prefix}ContactPoa`, result["fullname"], result["contactid"], "contact", "fullname")
        apiField(`${prefix}ContactLegal`, result["fullname"], result["contactid"], "contact", "fullname")
        if (result["ax_dnacompleted@OData.Community.Display.V1.FormattedValue"] == "Yes") {
            document.getElementById(`${prefix}DnaTable`).style.display = "block";
            document.getElementById(`${prefix}DnaNotFound`).style.display = "none";
            document.getElementById(`${prefix}FinaTable`).style.display = "none";
            apiField(`${prefix}RiskScore`, result["ax_riskbehaviourscore"], result["contactid"], "contact", "ax_riskbehaviourscore")
            apiField(`${prefix}PortRiskGroup`, result["ax_portfolioriskgroup"], result["contactid"], "contact", "ax_portfolioriskgroup")
            apiField(`${prefix}FinRelMgt`, result["ax_financialrelationshipmanagement"], result["contactid"], "contact", "ax_financialrelationshipmanagement")
            apiField(`${prefix}FinPlnMgt`, result["ax_financialplanningmanagement"], result["contactid"], "contact", "ax_financialplanningmanagement")
            apiField(`${prefix}WlthBldMtv`, result["ax_wealthbuildingmotivation"], result["contactid"], "contact", "ax_wealthbuildingmotivation")
            apiField(`${prefix}FinEmInt`, result["ax_financialemotionalintelligence"], result["contactid"], "contact", "ax_financialemotionalintelligence")
            apiField(`${prefix}NtBehSty`, result["ax_naturalbehaviourstyle@OData.Community.Display.V1.FormattedValue"], result["contactid"], "contact", "ax_naturalbehaviourstyle")
            apiField(`${prefix}Bias1`, result["ax_bias1@OData.Community.Display.V1.FormattedValue"], result["contactid"], "contact", "ax_bias1")
            apiField(`${prefix}Bias2`, result["ax_bias2@OData.Community.Display.V1.FormattedValue"], result["contactid"], "contact", "ax_bias2")
            apiField(`${prefix}PoaInPlace`, result["ax_poainplace@OData.Community.Display.V1.FormattedValue"], result["contactid"], "contact", "ax_poainplace")
        } else if (result["ax_finametricacompleted@OData.Community.Display.V1.FormattedValue"] == "Yes") {
            document.getElementById(`${prefix}DnaTable`).style.display = "none";
            document.getElementById(`${prefix}DnaNotFound`).style.display = "none";
            document.getElementById(`${prefix}FinaTable`).style.display = "block";
            apiField(`${prefix}FinaDateCom`, formattedDataFinaCompleted, result["contactid"], "contact", "ax_finametricacompleted")
            apiField(`${prefix}FinaScore`, result["ax_finametricascore"], result["contactid"], "contact", "ax_finametricascore")
            apiField(`${prefix}FinaRiskGroup`, result["ax_risk_group"], result["contactid"], "contact", "ax_risk_group")

        } else {
            document.getElementById(`${prefix}DnaTable`).style.display = "none";
            document.getElementById(`${prefix}DnaNotFound`).style.display = "block";
            document.getElementById(`${prefix}FinaTable`).style.display = "none";
        }
        apiField(`${prefix}PoaInPlace`, result["ax_poainplace@OData.Community.Display.V1.FormattedValue"], result["contactid"], "contact", "ax_poainplace")
            if (result["ax_poainplace@OData.Community.Display.V1.FormattedValue"] == "Yes") {
                apiField(`${prefix}PoaType`, result["ax_type@OData.Community.Display.V1.FormattedValue"], result["contactid"], "contact", "ax_type")
                apiField(`${prefix}PoaOnFile`, result["ax_poaonfile@OData.Community.Display.V1.FormattedValue"], result["contactid"], "contact", "ax_poaonfile")
                apiField(`${prefix}PoaStoredByEQ`, result["ax_poastoredbyeq@OData.Community.Display.V1.FormattedValue"], result["contactid"], "contact", "ax_poastoredbyeq")
                apiField(`${prefix}WhereStored`, result["ax_whereispoastored"], result["contactid"], "contact", "ax_whereispoastored")
                apiField(`${prefix}Registered`, result["ax_poaregisteredwithopg@OData.Community.Display.V1.FormattedValue"], result["contactid"], "contact", "ax_poaregisteredwithopg")
                apiField(`${prefix}PoainUse`, result["ax_poainuse@OData.Community.Display.V1.FormattedValue"], result["contactid"], "contact", "ax_poainuse")
                apiField(`${prefix}AttorneyNotes`, result["ax_attorneynotes"], result["contactid"], "contact", "ax_attorneynotes")
            } else {
                var elements = document.querySelectorAll(`#${prefix}PoaDetails`);
                elements.forEach(function(element) {
                element.style.display = 'none';
                })
            }
        apiField(`${prefix}WillInPlace`, result["ax_will_in_place@OData.Community.Display.V1.FormattedValue"], result["contactid"], "contact", "ax_will_in_place")
            if (result["ax_will_in_place@OData.Community.Display.V1.FormattedValue"] == "Yes") {
                apiField(`${prefix}WillOnFile`, result["ax_willonfile@OData.Community.Display.V1.FormattedValue"], result["contactid"], "contact", "ax_willonfile")
                apiField(`${prefix}WillStoredByEQ`, result["ax_willstoredbyeq@OData.Community.Display.V1.FormattedValue"], result["contactid"], "contact", "ax_willstoredbyeq")
                apiField(`${prefix}WhereWillStored`, result["ax_whereiswillstored"], result["contactid"], "contact", "ax_whereiswillstored")
                apiField(`${prefix}DateOfWill`, formattedDateOfWill, result["contactid"], "contact", "ax_dateofwill")
                apiField(`${prefix}AssetsIntoTrust`, result["ax_brassetsintotruston1stdeath@OData.Community.Display.V1.FormattedValue"], result["contactid"], "contact", "ax_brassetsintotruston1stdeath")
            } else {
                var elements = document.querySelectorAll(`#${prefix}WillDetails`);
                elements.forEach(function(element) {
                element.style.display = 'none';
                })
            }
        apiField(`${prefix}KnownAsTax`, result["fullname"], result["contactid"], "contact", "fullname")
        apiField(`${prefix}IsaContributions`, result["ax_isaallowanceused@OData.Community.Display.V1.FormattedValue"], result["contactid"], "contact", "ax_isaallowanceused")
        apiField(`${prefix}PensionContributions`, result["ax_pensionallowanceused2425@OData.Community.Display.V1.FormattedValue"], result["contactid"], "contact", "ax_pensionallowanceused2425")
        apiField(`${prefix}Cgt`, result["ax_cgt2425@OData.Community.Display.V1.FormattedValue"], result["contactid"], "contact", "ax_cgt2425")
        apiField(`${prefix}MarriageAllowance`, result["ax_marriageallowanceused@OData.Community.Display.V1.FormattedValue"], result["contactid"], "contact", "ax_marriageallowanceused")
        if (result["ax_marriageallowanceused@OData.Community.Display.V1.FormattedValue"] == "Yes") {
            apiField(`${prefix}MarriageAllowanceNotes`, result["ax_marriageallowancenotes"], result["contactid"], "contact", "ax_marriageallowancenotes")

        } else {
            var elements = document.querySelectorAll(`#${prefix}MarrAllow`);
                elements.forEach(function(element) {
                element.style.display = 'none';
                })
        }
        apiField(`${prefix}PensionProtection`, result["ax_protectioninplace@OData.Community.Display.V1.FormattedValue"], result["contactid"], "contact", "ax_protectioninplace")
        if ( result["ax_protectioninplace@OData.Community.Display.V1.FormattedValue"] == "Yes") {
            apiField(`${prefix}CertOnFile`, result["ax_certificateonfile@OData.Community.Display.V1.FormattedValue"], result["contactid"], "contact", "ax_certificateonfile")
            apiField(`${prefix}ProtLTAAmount`, result["ax_protectedltaamount@OData.Community.Display.V1.FormattedValue"], result["contactid"], "contact", "ax_protectedltaamount")
            apiField(`${prefix}PensionProtectionType`, result["_ax_protectiontype_value@OData.Community.Display.V1.FormattedValue"], result["contactid"], "contact", "_ax_protectiontype_value")
        } else {
            var elements = document.querySelectorAll(`#${prefix}PensionProt`);
                elements.forEach(function(element) {
                element.style.display = 'none';
                })
        }
   
   
   
    });
}

//Starts the spinning of the loading icon
function startSpinning() {
    const spinner = document.getElementById('spinner');
    spinner.style.animationPlayState = 'running';
}

//Stops the spinnign of the loading icon
function stopSpinning() {
    const spinner = document.getElementById('spinner');
    spinner.style.animationPlayState = 'paused';
}

function firstLoad(elementId, textId, tabName) {
    // Show the loading icon
    document.getElementById(tabName).style.display = "none";
    document.getElementById(elementId).style.display = "block";
    //---------------------
    //Declaring Lookup Fields
    const lookupFields = [
        { wrField: "financialPlanner", crmField: "ax_financialplanner"},
        { wrField: "solicitor", crmField: "ax_solicitor" },
        { wrField: "accountant", crmField: "ax_accountant" }

    ]
    //Setting Lookup Fields
    lookupFields.forEach(field => {
        getLookupFieldName(field.wrField, field.crmField);
    });
    //---------------------
    //Declaring Years and Days Fields
    const yearsAndDaysFields = [
        { wrField: "fpServiceYears", crmYearsField: "ax_fpyearsservice", crmDaysField: "ax_fpyearsserviceplusdays" },
        { wrField: "clientYears", crmYearsField: "ax_yearsaclient", crmDaysField: "ax_yearsaclientplusdays" }
    ]
    //Setting Years and Days Fields
    yearsAndDaysFields.forEach(field => {
        getYearsAndDays(field.wrField, field.crmYearsField, field.crmDaysField);
    });
    //---------------------
    //Declaring Optionset Fields
    const optionsetFields = [
        { wrField: "serviceProvided", crmField: "ax_serviceprovided"},
        { wrField: "primSeg", crmField: "ax_primarysegment"},
        { wrField: "ageSeg", crmField: "ax_agesegment"},
        { wrField: "colourCode", crmField: "ax_colourcode"},
        { wrField: "lossRating", crmField: "ax_capacityrating" }
    ]
    //Setting Optionset Fields
    optionsetFields.forEach(field => {
        getOptionsetFieldName(field.wrField, field.crmField);
    });
    //---------------------
    //Declaring Fields
    const fields = [
        { wrField: "clientRating", crmField: "ax_clientrating" },
        { wrField: "primObj", crmField: "ax_primaryobjective" },
        { wrField: "secObj", crmField: "ax_secondaryobjectives" },
        { wrField: "background", crmField: "ax_clientbackground" },
        { wrField: "knowledge", crmField: "ax_knowledgeexperience" },
        { wrField: "needForRisk", crmField: "ax_needforrisk" },
        { wrField: "desireForRisk", crmField: "ax_desireforrisk" },
        { wrField: "capacityForLoss", crmField: "ax_capacityforloss" }
    ]
    //Setting Fields
    fields.forEach(field => {
        getFieldValue(field.wrField, field.crmField);
    });    
    //---------------------
    //Declaring Rating Fields
    const ratingFields = [
        { wrField: "portfolio", crmField: "ax_portfolio" },
        { wrField: "financialPlan", crmField: "ax_financialplan" },
        { wrField: "relationship", crmField: "ax_relationship" }
    ]
    //Setting Rating Fields
    ratingFields.forEach(field => {
        getFieldValueRating(field.wrField, field.crmField);
    });
    //---------------------
    //Declaring Formatted Fields
    const formattedFields = [
        { wrField: "costOfLiving", crmField: "ax_costofliving" },
        { wrField: "discretionary", crmField: "ax_discretionaryexpenditure" },
        { wrField: "shortTerm", crmField: "ax_shorttermadditionalexpenditure" },
        { wrField: "totalEx", crmField: "ax_totalexpenditure" }
    ]
    //Setting Formatted Fields
    formattedFields.forEach(field => {
        getFormattedValueFieldName(field.wrField, field.crmField);
    });
    //---------------------
    //Get Account ID
    accountId = window.parent.Xrm.Page.data.entity.getId();
    //Declaring Table Values
    const tableValues = [
        //Income
        { 
            wrField: "employmentTable",
            filter: "_ax_account_value",
            entity: "ax_income",
            array: [
            { width: "150px", field: "ax_category@OData.Community.Display.V1.FormattedValue" },
            { width: "250px", field: "_ax_contact_value@OData.Community.Display.V1.FormattedValue" },
            { width: "350px", field: "ax_companyname" },
            { width: "150px", field: "ax_totalremuneration@OData.Community.Display.V1.FormattedValue" }],
            formGuid: "d9f5ee5b-7bd8-4c83-8bd8-b704f1c9da34"
        },
        //Cashflow
        {
            wrField: "cashTable",
            filter: "_ax_account_value",
            entity: "ax_asset",
            array: [
            { width: "250px", field: "ax_category@OData.Community.Display.V1.FormattedValue" },
            { width: "250px", field: "ax_assetowner" },
            { width: "350px", field: "ax_name" },
            { width: "100px", field: "ax_totalvaluation@OData.Community.Display.V1.FormattedValue" }],
            formGuid: "9bde7472-a0fe-eb11-94ef-0022481b45f8"
        },
        //Facts
        {
            wrField: "factTable",
            filter: "_ax_account_value",
            entity: "ax_keyfact",
            array: [
            { width: "250px", field: "_ax_contact_value@OData.Community.Display.V1.FormattedValue" },
            { width: "150px", field: "ax_significance@OData.Community.Display.V1.FormattedValue" },
            { width: "450px", field: "ax_comments" }],
            formGuid: ""
        },
        //Feedback
        {
            wrField: "feedbackTable",
            filter: "_ax_account_value",
            entity: "ax_feedback",
            array: [
            { width: "250px", field: "ax_category@OData.Community.Display.V1.FormattedValue" },
            { width: "150px", field: "ax_rating@OData.Community.Display.V1.FormattedValue" },
            { width: "450px", field: "ax_comments" }],
            formGuid: ""
        },
        //Our Actions
        {
            wrField: "ourActionTable",
            filter: "ax_type ne 100000001 and _ax_account_value",
            entity: "task",
            array: [
            { width: "600px", field: "description" },
            { width: "150px", field: "ax_duedate@OData.Community.Display.V1.FormattedValue" }],
            formGuid: ""
        },
        //Client Actions
        {
            wrField: "clientActionTable",
            filter: "ax_type eq 100000001 and _ax_account_value",
            entity: "task",
            array: [
            { width: "600px", field: "description" },
            { width: "150px", field: "ax_duedate@OData.Community.Display.V1.FormattedValue" }],
            formGuid: ""
        },
        //Trusts
        {
            wrField: "trustTable",
            filter: "_cr330_trust_account_value",
            entity: "ax_trust",
            array: [
                { width: "250px", field: "ax_trusttype@OData.Community.Display.V1.FormattedValue" },
                { width: "250px", field: "ax_truststartdate@OData.Community.Display.V1.FormattedValue"},
                { width: "350px", field: "ax_trust"},
                { width: "100px", field: "ax_ageoftrust"}
            ],
            formGuid: ""
        },
        //Companies
        {
            wrField: "companyTable",
            filter: "_cr330_company_account_value",
            entity: "ax_company",
            array: [
                { width: "250px", field: "ax_businesstype@OData.Community.Display.V1.FormattedValue" },
                { width: "250px", field: "ax_companynumber"},
                { width: "350px", field: "ax_name"},
                { width: "100px", field: "ax_nextyearend@OData.Community.Display.V1.FormattedValue"}
            ],
            formGuid: ""
        }
       
    ]
    //Setting Tables
    tableValues.forEach(table => {
        createTableOfEntityLinkedToEntity(table.wrField, table.filter, accountId, table.entity, table.array, table.formGuid)
    })
    //Filter Tables after generation
    tableFilter("incomeType", "employmentTable")
    tableFilter("assetType" , "cashTable");
    //---------------------
    //Get Date Became Client
    if (window.parent.Xrm.Page.getAttribute("ax_datebecameclient").getValue()) {
            document.getElementById("dateBecameClient").textContent = window.parent.Xrm.Page.getAttribute("ax_datebecameclient").getValue().toLocaleDateString('en-GB', {
                day: '2-digit',
                month: '2-digit',
                year: 'numeric'
            });
    } else {
            tabFocusLocal("ax_datebecameclient", "dateBecameClient");
    }
    //---------------------
    //Set Primary Contact Fields
    setContactFields("primarycontactid", "pc");
    //Set Secondary Contact Fields
    if (window.parent.Xrm.Page.getAttribute("ax_secondarycontact").getValue() !== null) {
            setContactFields("ax_secondarycontact", "sc");
            familyValues = [
                { width: "250px", field: "_record2id_value@OData.Community.Display.V1.FormattedValue" },
                { width: "150px", field: "_record2roleid_value@OData.Community.Display.V1.FormattedValue" },
                { width: "150px", field: "status" },
                { width: "150px", field: "age" }
            ];
    
            getFamilyRoleGuids(familyValues, "sc", "ax_secondarycontact");
    } else {
            const secondContactFields = [ {wrField: "secConOverview"}, {wrField: "secConDNA"}, {wrField: "secConWills"}, {wrField: "secConPoa"}, {wrField: "secConTax"} ]
            secondContactFields.forEach(field => {
                document.getElementById(field.wrField).style.display = "none";
            })
    }
    //---------------------
    //Primary Contact Family Tree
    familyValues = [
            { width: "250px", field: "_record2id_value@OData.Community.Display.V1.FormattedValue" },
            { width: "150px", field: "_record2roleid_value@OData.Community.Display.V1.FormattedValue" },
            { width: "150px", field: "status" },
            { width: "150px", field: "age" }
    ];
    //Set Family Rows Clickable
    getFamilyRoleGuids(familyValues, "pc", "primarycontactid");
    //---------------------
    // Hide the loading icon
    setTimeout(() => {
        document.getElementById(tabName).style.display = "inline";
        document.getElementById(elementId).style.display = "none";
    }, 2000); 
}

//Gets the contact entity
function getContact(guid) {
    return window.parent.Xrm.WebApi.retrieveRecord("contact", guid);
}

function getFamilyRoleGuids(familyValues, prefix, primaryOrSecondary) {
    var query = `?$filter=category eq 2 and statecode eq 0`;
    var resultString = "";
    window.parent.Xrm.WebApi.retrieveMultipleRecords("connectionrole", query).then(
        function (response) {
            var entities = response.entities;

            entities.forEach(function(entity, index) {
                resultString += "_record2roleid_value eq "
                resultString += entity["connectionroleid"]
                if (index < entities.length - 1) {
                    resultString += " or "
                }
            })

            getContact(window.parent.Xrm.Page.getAttribute(`${primaryOrSecondary}`).getValue()[0].id).then(function(result) {
                var contactId = result["contactid"];
                createTableOfEntityLinkedToEntity(`${prefix}FamilyTree`, `(${resultString}) and _record1id_value`, contactId, "connection", familyValues);
            })
        }
    )
}

//On load of tab trigger initial functions
document.addEventListener("DOMContentLoaded", () => setInitialFocus());
