function toggleDiv(divName, activeDiv) {
    try { 
        document.getElementById(activeDiv).style.display = "none";
        document.getElementById(divName).style.display = "inline";
    } catch (error) {
        console.error("Error in toggleDiv:", error);
    }
}

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

function transformSync() {
    document.getElementById("syncIcon").style.transform = 'rotate(360deg)';
}

function resetSync() {
    document.getElementById("syncIcon").style.transform = 'rotate(0deg)';
}

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
    ]

    elementIds.forEach(function(element){
        if (element.id == "loading") {
            document.getElementById(element.id).style.display = "inline";
        } else {
            document.getElementById(element.id).style.display = "none";
        }
    })

    document.getElementById("syncButton").addEventListener('mouseover', transformSync);
    document.getElementById("syncButton").addEventListener('mouseout', resetSync);
    firstLoad("loading", "loadText");
}

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

function createTableOfEntityLinkedToEntity(tableId, fieldName, recordId, entityName, rows, formId) {
    var query = `?$filter=${fieldName} eq ${encodeURIComponent("'" + recordId + "'")} and statecode eq 0`;

    window.parent.Xrm.WebApi.retrieveMultipleRecords(entityName, query).then(
        function (response) {
            var entities = response.entities;

            // Get a reference to the table
            var table = document.getElementById(tableId);

            // Clear existing rows (optional)
            while (table.rows.length > 1) {
                table.deleteRow(1);
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
                }

                newRow.addEventListener('click', function(event) {
                    event.preventDefault();
                    openFormWithTab(entityName, rowId, getTabNameOfField(entityName, formId, income[rows[3].field]));
                })

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
        target: 2
    }

    window.parent.Xrm.Navigation.navigateTo(pageInput, navigationOptions).then(
        function (success) {
                    console.log("Form opened successfully.");
                    document.getElementById("loading").style.display = "none";
                },
                function (error) {
                    console.error("Error opening form: " + error.message);
                }
            );
}

function jumpToLocalTab(tabName) {
    var formContext = window.parent.Xrm.Page;
    var tab = formContext.ui.tabs.get(tabName);
    
    if (tab) {
        tab.setFocus();
    } else {
        console.error("Tab not found: " + tabName);
    }
}

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

function getLookupFieldName(elementId, fieldName) {
    if (window.parent.Xrm.Page.getAttribute(fieldName).getValue()) {
        document.getElementById(elementId).textContent =  window.parent.Xrm.Page.getAttribute(fieldName).getValue()[0].name;
    } else {
        tabFocusLocal(fieldName, elementId)
    }
}

function getOptionsetFieldName(elementId, fieldName) {
    if (window.parent.Xrm.Page.getAttribute(fieldName).getSelectedOption()) {
        document.getElementById(elementId).textContent =  window.parent.Xrm.Page.getAttribute(fieldName).getSelectedOption().text
    } else {
        tabFocusLocal(fieldName, elementId)
    }
}

function formatCurrency(value, locale = 'en-GB', currency = 'GBP') {
    return new Intl.NumberFormat(locale, { style: 'currency', currency: currency }).format(value);
}

function getFormattedValueFieldName(elementId, fieldName) {
    if (window.parent.Xrm.Page.getAttribute(fieldName).getValue()) {
        var amount = window.parent.Xrm.Page.getAttribute(fieldName).getValue()
        var formattedAmount = formatCurrency(amount);
        document.getElementById(elementId).textContent =  formattedAmount;
    } else {
        tabFocusLocal(fieldName, elementId)
    }
}

function getYearsAndDays(elementId, fieldName1, fieldName2) {
    if (window.parent.Xrm.Page.getAttribute(fieldName2).getValue()) {
        document.getElementById(elementId).textContent =  `${window.parent.Xrm.Page.getAttribute(fieldName1).getValue()} years and ${window.parent.Xrm.Page.getAttribute(fieldName2).getValue()} days`
    } else {
        tabFocusLocal(fieldName1, elementId)
    }
}

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

        apiField(`${prefix}KnownAs`, result["ax_knownas"], result["contactid"], "contact", "ax_knownas")
        apiField(`${prefix}Dob`, formattedDateOfBirth, result["contactid"], "contact", "ax_dateofbirth")
        apiField(`${prefix}Employment`, result["ax_employmentstatus@OData.Community.Display.V1.FormattedValue"], result["contactid"], "contact", "ax_employmentstatus")
        apiField(`${prefix}Nationality`, result["ax_nationality"], result["contactid"], "contact","ax_nationality")
        apiField(`${prefix}Residency`, result["ax_residency@OData.Community.Display.V1.FormattedValue"] , result["contactid"], "contact", "ax_residency")
        apiField(`${prefix}Domicile`, result["ax_domicile@OData.Community.Display.V1.FormattedValue"], result["contactid"], "contact", "ax_domicile")
        apiField(`${prefix}Health`, result["ax_healthcomments"], result["contactid"], "contact", "ax_healthcomments")
        apiField(`${prefix}Name`, result["fullname"], result["contactid"], "contact", "fullname")
        apiField(`${prefix}ContactPoa`, result["fullname"], result["contactid"], "contact", "fullname")
        apiField(`${prefix}ContactLegal`, result["fullname"], result["contactid"], "contact", "fullname")
        if (result["ax_dnacompleted@OData.Community.Display.V1.FormattedValue"] == "Yes") {
            document.getElementById(`${prefix}DnaTable`).style.display = "block";
            document.getElementById(`${prefix}DnaNotFound`).style.display = "none";
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
        } else {
            document.getElementById(`${prefix}DnaTable`).style.display = "none";
            document.getElementById(`${prefix}DnaNotFound`).style.display = "block";
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
    });
}

function startSpinning() {
    const spinner = document.getElementById('spinner');
    spinner.style.animationPlayState = 'running';
}

function stopSpinning() {
    const spinner = document.getElementById('spinner');
    spinner.style.animationPlayState = 'paused';
}

function firstLoad(elementId, textId) {
    // Show the loading icon
    document.getElementById("overview").style.display = "none";
    document.getElementById(elementId).style.display = "block";

        getLookupFieldName("financialPlanner", "ax_financialplanner");
        getYearsAndDays("fpServiceYears", "ax_fpyearsservice", "ax_fpyearsserviceplusdays");
        if (window.parent.Xrm.Page.getAttribute("ax_datebecameclient").getValue()) {
            document.getElementById("dateBecameClient").textContent = window.parent.Xrm.Page.getAttribute("ax_datebecameclient").getValue().toLocaleDateString('en-GB', {
                day: '2-digit',
                month: '2-digit',
                year: 'numeric'
            });
        } else {
            tabFocusLocal("ax_datebecameclient", "dateBecameClient");
        }
        getYearsAndDays("clientYears", "ax_yearsaclient", "ax_yearsaclientplusdays");
        getOptionsetFieldName("serviceProvided", "ax_serviceprovided");
        getOptionsetFieldName("primSeg", "ax_primarysegment");
        getOptionsetFieldName("ageSeg", "ax_agesegment");
        getOptionsetFieldName("colourCode", "ax_colourcode");
        getFieldValue("clientRating", "ax_clientrating");
        getFieldValueRating("portfolio", "ax_portfolio");
        getFieldValueRating("financialPlan", "ax_financialplan");
        getFieldValueRating("relationship", "ax_relationship");
        setContactFields("primarycontactid", "pc");
        if (window.parent.Xrm.Page.getAttribute("ax_secondarycontact").getValue() !== null) {
            setContactFields("ax_secondarycontact", "sc");
        } else {
            document.getElementById("secConOverview").style.display = "none";
            document.getElementById("secConDNA").style.display = "none";
            document.getElementById("secConWills").style.display = "none";
            document.getElementById("secConPoa").style.display = "none";
        }
        getFieldValue("primObj", "ax_primaryobjective");
        getFieldValue("secObj", "ax_secondaryobjectives");
        getFieldValue("background", "ax_clientbackground")
        getFieldValue("knowledge", "ax_knowledgeexperience");
        getFieldValue("needForRisk", "ax_needforrisk");
        getFieldValue("desireForRisk", "ax_desireforrisk");
        getOptionsetFieldName("lossRating", "ax_capacityrating");
        getFieldValue("capacityForLoss", "ax_capacityforloss");
        accountId = window.parent.Xrm.Page.data.entity.getId();
        employmentValues = [
            { width: "150px", field: "ax_category@OData.Community.Display.V1.FormattedValue" },
            { width: "250px", field: "_ax_contact_value@OData.Community.Display.V1.FormattedValue" },
            { width: "350px", field: "ax_companyname" },
            { width: "150px", field: "ax_totalremuneration@OData.Community.Display.V1.FormattedValue" }
        ];
        createTableOfEntityLinkedToEntity("employmentTable", "_ax_account_value", accountId, "ax_income", employmentValues, "d9f5ee5b-7bd8-4c83-8bd8-b704f1c9da34");
        getFormattedValueFieldName("costOfLiving", "ax_costofliving");
        getFormattedValueFieldName("discretionary", "ax_discretionaryexpenditure");
        getFormattedValueFieldName("shortTerm", "ax_shorttermadditionalexpenditure");
        getFormattedValueFieldName("totalEx", "ax_totalexpenditure");
        cashValues = [
            { width: "250px", field: "ax_category@OData.Community.Display.V1.FormattedValue" },
            { width: "250px", field: "ax_assetowner" },
            { width: "350px", field: "ax_name" },
            { width: "100px", field: "ax_totalvaluation@OData.Community.Display.V1.FormattedValue" }
        ];
        createTableOfEntityLinkedToEntity("cashTable", "_ax_account_value", accountId, "ax_asset", cashValues, "9bde7472-a0fe-eb11-94ef-0022481b45f8");
        tableFilter("incomeType", "employmentTable")
        tableFilter("assetType" , "cashTable");

        getLookupFieldName("solicitor", "ax_solicitor");
        getLookupFieldName("accountant", "ax_accountant");

        //Facts
        factValues = [
            { width: "250px", field: "_ax_contact_value@OData.Community.Display.V1.FormattedValue" },
            { width: "150px", field: "ax_significance@OData.Community.Display.V1.FormattedValue" },
            { width: "450px", field: "ax_comments" }
        ];
        createTableOfEntityLinkedToEntity("factTable", "_ax_account_value", accountId, "ax_keyfact", factValues);

        //Feedback
        feedbackValues = [
            { width: "250px", field: "ax_category@OData.Community.Display.V1.FormattedValue" },
            { width: "150px", field: "ax_rating@OData.Community.Display.V1.FormattedValue" },
            { width: "450px", field: "ax_comments" }
        ];
        createTableOfEntityLinkedToEntity("feedbackTable", "_ax_account_value", accountId, "ax_feedback", feedbackValues);

        //Our Actions
        ourActionsValues = [
            { width: "600px", field: "description" },
            { width: "150px", field: "ax_duedate@OData.Community.Display.V1.FormattedValue" }
        ];
        createTableOfEntityLinkedToEntity("ourActionTable", "ax_type ne 100000001 and _regardingobjectid_value", accountId, "task", ourActionsValues);

        //Client Actions
        clientActionsValues = [
            { width: "600px", field: "description" },
            { width: "150px", field: "ax_duedate@OData.Community.Display.V1.FormattedValue" }
        ];
        createTableOfEntityLinkedToEntity("clientActionTable", "ax_type eq 100000001 and _regardingobjectid_value", accountId, "task", clientActionsValues);

        // Hide the loading icon
    setTimeout(() => {
        document.getElementById("overview").style.display = "inline";
        document.getElementById(elementId).style.display = "none";
    }, 2000); 
}

function getContact(guid) {
    return window.parent.Xrm.WebApi.retrieveRecord("contact", guid);
}

document.addEventListener("DOMContentLoaded", () => setInitialFocus());