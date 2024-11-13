function toggleDiv(divName, activeDiv) {
    try { 
        const divToHide = document.getElementById(getActiveDiv(activeDiv));
        if (divToHide) {
            divToHide.style.display = "none";
        }

        let x;
        switch (divName) {
            case "overview":
                x = document.getElementById("overview");
                break;
            case "objectives":
                x = document.getElementById("objectives");
                break;
            case "risk":
                x = document.getElementById("risk");
                break;
            case "dna":
                x = document.getElementById("dna");
                break;
            case "income":
                x = document.getElementById("income");
                break;
            case "asset":
                x = document.getElementById("asset");
                break;
            case "prof":
                x = document.getElementById("prof");
                break;
            case "fnf":
                x = document.getElementById("fnf");
                break;
            case "actions":
                x = document.getElementById("actions");
                break;
            default:
                x = document.getElementById("overview");
        }

        if (x) {
            x.style.display = "inline";
        }
    } catch (error) {
        console.error("Error in toggleDiv:", error);
    }
}

function incomeFilter() {
    document.getElementById('incomeType').addEventListener('change', function() {
        var filter = this.value;
        var rows = document.querySelectorAll('#employmentTable tbody tr');
      
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

function assetFilter() {
    document.getElementById('assetType').addEventListener('change', function() {
        var filter = this.value;
        var rows = document.querySelectorAll('#cashTable tbody tr');
      
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

function getActiveDiv(activeDiv) {
    switch (activeDiv) {
        case "overview":
            return "overview";
        case "objectives":
            return "objectives";
        case "risk":
            return "risk"
        case "dna":
            return "dna"
        case "income":
            return "income"
        case "asset":
            return "asset"
        case "prof":
            return "prof"
        case "fnf":
            return "fnf"
        case "actions":
            return "actions"
        default:
            throw new Error("Invalid activeDiv");
    }
}

function setInitialFocus() {
    document.getElementById("loading").style.display = "inline";
    document.getElementById("overview").style.display = "none";
    document.getElementById("objectives").style.display = "none";
    document.getElementById("risk").style.display = "none";
    document.getElementById("dna").style.display = "none";
    document.getElementById("income").style.display = "none";
    document.getElementById("asset").style.display = "none";
    document.getElementById("prof").style.display = "none";
    document.getElementById("fnf").style.display = "none";
    document.getElementById("actions").style.display = "none";
    document.getElementById("syncing").style.display = "none";
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

function createTableOfEntityLinkedToEntity(tableId, fieldName, recordId, entityName, rows) {
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

            if (entityName == "js_income") {
                sortedData = entities.sort((a, b) => {
                    if (a["js_category@OData.Community.Display.V1.FormattedValue"] < b["js_category@OData.Community.Display.V1.FormattedValue"]) {
                      return -1;
                    }
                    if (a["js_category@OData.Community.Display.V1.FormattedValue"] > b["js_category@OData.Community.Display.V1.FormattedValue"]) {
                      return 1;
                    }
                    // If js_category is the same, sort by js_contact_value
                    if (a["_js_contact_value@OData.Community.Display.V1.FormattedValue"] < b["_js_contact_value@OData.Community.Display.V1.FormattedValue"]) {
                      return -1;
                    }
                    if (a["_js_contact_value@OData.Community.Display.V1.FormattedValue"] > b["_js_contact_value@OData.Community.Display.V1.FormattedValue"]) {
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

                rows.forEach(function(row, cellIndex) {
                    var newCell = newRow.insertCell(cellIndex);
                    newCell.style.width = row.width;
                    switch (entityName) {
                      case "js_income":
                        switch (row.field) {
                          case "js_totalremuneration@OData.Community.Display.V1.FormattedValue":
                            if (income[rows[0].field] == "Employment") {
                              newCell.textContent = income[row.field] || "-";
                            } else {
                              newCell.textContent = income["js_annual_amount@OData.Community.Display.V1.FormattedValue"] || "-";
                            }
                            break;
                          case "js_companyname":
                            switch (income[rows[0].field]) {
                              case "Employment":
                                newCell.textContent = income[row.field] || "-";
                                break;
                              case "Asset based":
                                newCell.textContent = income["_js_asset_value@OData.Community.Display.V1.FormattedValue"] || "-";
                                break;
                              case "Other":
                                newCell.textContent = income["js_source"] || "-";
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
        tab = getTabNameOfField(recordType, "guid", fieldName).then(
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
        var dateOfBirth = result["js_dateofbirth"];
        var dateOfBirthClean = new Date(dateOfBirth);
        var formattedDateOfBirth = dateOfBirthClean.toLocaleString('en-GB', {
            day: '2-digit',
            month: '2-digit',
            year: 'numeric'
        })

        var dateOfWill = result["js_dateofwill"];
        var dateOfWillClean = new Date(dateOfWill);
        var formattedDateOfWill = dateOfWillClean.toLocaleString('en-GB', {
            day: '2-digit',
            month: '2-digit',
            year: 'numeric'
        })
        

        apiField(`${prefix}KnownAs`, result["js_knownas"], result["contactid"], "contact", "js_knownas")
        apiField(`${prefix}Dob`, formattedDateOfBirth, result["contactid"], "contact", "js_dateofbirth")
        apiField(`${prefix}Employment`, result["js_employmentstatus@OData.Community.Display.V1.FormattedValue"], result["contactid"], "contact", "js_employmentstatus")
        apiField(`${prefix}Nationality`, result["js_nationality"], result["contactid"], "contact","js_nationality")
        apiField(`${prefix}Residency`, result["js_residency@OData.Community.Display.V1.FormattedValue"] , result["contactid"], "contact", "js_residency")
        apiField(`${prefix}Domicile`, result["js_domicile@OData.Community.Display.V1.FormattedValue"], result["contactid"], "contact", "js_domicile")
        apiField(`${prefix}Health`, result["js_healthcomments"], result["contactid"], "contact", "js_healthcomments")
        apiField(`${prefix}Name`, result["fullname"], result["contactid"], "contact", "fullname")
        apiField(`${prefix}ContactLegal`, result["fullname"], result["contactid"], "contact", "fullname")
        apiField(`${prefix}RiskScore`, result["js_riskbehaviourscore"], result["contactid"], "contact", "js_riskbehaviourscore")
        apiField(`${prefix}PortRiskGroup`, result["js_portfolioriskgroup"], result["contactid"], "contact", "js_portfolioriskgroup")
        apiField(`${prefix}FinRelMgt`, result["js_financialrelationshipmanagement"], result["contactid"], "contact", "js_financialrelationshipmanagement")
        apiField(`${prefix}FinPlnMgt`, result["js_financialplanningmanagement"], result["contactid"], "contact", "js_financialplanningmanagement")
        apiField(`${prefix}WlthBldMtv`, result["js_wealthbuildingmotivation"], result["contactid"], "contact", "js_wealthbuildingmotivation")
        apiField(`${prefix}FinEmInt`, result["js_financialemotionalintelligence"], result["contactid"], "contact", "js_financialemotionalintelligence")
        apiField(`${prefix}NtBehSty`, result["js_naturalbehaviourstyle@OData.Community.Display.V1.FormattedValue"], result["contactid"], "contact", "js_naturalbehaviourstyle")
        apiField(`${prefix}Bias1`, result["js_bias1@OData.Community.Display.V1.FormattedValue"], result["contactid"], "contact", "js_bias1")
        apiField(`${prefix}Bias2`, result["js_bias2@OData.Community.Display.V1.FormattedValue"], result["contactid"], "contact", "js_bias2")
        apiField(`${prefix}PoaInPlace`, result["js_poainplace@OData.Community.Display.V1.FormattedValue"], result["contactid"], "contact", "js_poainplace")
            if (result["js_poainplace@OData.Community.Display.V1.FormattedValue"] == "Yes") {
                apiField(`${prefix}PoaType`, result["js_type@OData.Community.Display.V1.FormattedValue"], result["contactid"], "contact", "js_type")
                apiField(`${prefix}PoaOnFile`, result["js_poaonfile@OData.Community.Display.V1.FormattedValue"], result["contactid"], "contact", "js_poaonfile")
                apiField(`${prefix}PoaStoredByEQ`, result["js_poastoredbyeq@OData.Community.Display.V1.FormattedValue"], result["contactid"], "contact", "js_poastoredbyeq")
                apiField(`${prefix}WhereStored`, result["js_whereispoastored"], result["contactid"], "contact", "js_whereispoastored")
                apiField(`${prefix}Registered`, result["js_poaregisteredwithopg@OData.Community.Display.V1.FormattedValue"], result["contactid"], "contact", "js_poaregisteredwithopg")
                apiField(`${prefix}PoainUse`, result["js_poainuse@OData.Community.Display.V1.FormattedValue"], result["contactid"], "contact", "js_poainuse")
                apiField(`${prefix}AttorneyNotes`, result["js_attorneynotes"], result["contactid"], "contact", "js_attorneynotes")
            } else {
                var elements = document.querySelectorAll(`#${prefix}PoaDetails`);
                elements.forEach(function(element) {
                element.style.display = 'none';
                })
            }
        apiField(`${prefix}WillInPlace`, result["js_will_in_place@OData.Community.Display.V1.FormattedValue"], result["contactid"], "contact", "js_will_in_place")
            if (result["js_will_in_place@OData.Community.Display.V1.FormattedValue"] == "Yes") {
                apiField(`${prefix}WillOnFile`, result["js_willonfile@OData.Community.Display.V1.FormattedValue"], result["contactid"], "contact", "js_willonfile")
                apiField(`${prefix}WillStoredByEQ`, result["js_willstoredbyeq@OData.Community.Display.V1.FormattedValue"], result["contactid"], "contact", "js_willstoredbyeq")
                apiField(`${prefix}WhereWillStored`, result["js_whereiswillstored"], result["contactid"], "contact", "js_whereiswillstored")
                apiField(`${prefix}DateOfWill`, formattedDateOfWill, result["contactid"], "contact", "js_dateofwill")
                apiField(`${prefix}AssetsIntoTrust`, result["js_brassetsintotruston1stdeath@OData.Community.Display.V1.FormattedValue"], result["contactid"], "contact", "js_brassetsintotruston1stdeath")
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

        getLookupFieldName("financialPlanner", "js_financialplanner");
        getYearsAndDays("fpServiceYears", "js_fpyearsservice", "js_fpyearsserviceplusdays");
        if (window.parent.Xrm.Page.getAttribute("js_datebecameclient").getValue()) {
            document.getElementById("dateBecameClient").textContent = window.parent.Xrm.Page.getAttribute("js_datebecameclient").getValue().toLocaleDateString('en-GB', {
                day: '2-digit',
                month: '2-digit',
                year: 'numeric'
            });
        } else {
            tabFocusLocal("js_datebecameclient", "dateBecameClient");
        }
        getYearsAndDays("clientYears", "js_yearsaclient", "js_yearsaclientplusdays");
        getOptionsetFieldName("serviceProvided", "js_serviceprovided");
        getOptionsetFieldName("primSeg", "js_primarysegment");
        getOptionsetFieldName("ageSeg", "js_agesegment");
        getOptionsetFieldName("colourCode", "js_colourcode");
        getFieldValue("clientRating", "js_clientrating");
        getFieldValueRating("portfolio", "js_portfolio");
        getFieldValueRating("financialPlan", "js_financialplan");
        getFieldValueRating("relationship", "js_relationship");
        setContactFields("primarycontactid", "pc");
        if (window.parent.Xrm.Page.getAttribute("js_secondarycontact").getValue() !== null) {
            setContactFields("js_secondarycontact", "sc");
        } else {
            document.getElementById("secConOverview").style.display = "none";
            document.getElementById("secConDNA").style.display = "none";
        }
        getFieldValue("primObj", "js_primaryobjective");
        getFieldValue("secObj", "js_secondaryobjectives");
        getFieldValue("knowledge", "js_knowledgeexperience");
        getFieldValue("needForRisk", "js_needforrisk");
        getFieldValue("desireForRisk", "js_desireforrisk");
        getOptionsetFieldName("lossRating", "js_capacityrating");
        getFieldValue("capacityForLoss", "js_capacityforloss");
        accountId = window.parent.Xrm.Page.data.entity.getId();
        employmentValues = [
            { width: "150px", field: "js_category@OData.Community.Display.V1.FormattedValue" },
            { width: "250px", field: "_js_contact_value@OData.Community.Display.V1.FormattedValue" },
            { width: "350px", field: "js_companyname" },
            { width: "150px", field: "js_totalremuneration@OData.Community.Display.V1.FormattedValue" }
        ];
        createTableOfEntityLinkedToEntity("employmentTable", "_js_account_value", accountId, "js_income", employmentValues);
        getFormattedValueFieldName("costOfLiving", "js_costofliving");
        getFormattedValueFieldName("discretionary", "js_discretionaryexpenditure");
        getFormattedValueFieldName("shortTerm", "js_shorttermadditionalexpenditure");
        getFormattedValueFieldName("totalEx", "js_totalexpenditure");
        cashValues = [
            { width: "250px", field: "js_category@OData.Community.Display.V1.FormattedValue" },
            { width: "250px", field: "js_assetowner" },
            { width: "350px", field: "js_name" },
            { width: "100px", field: "js_totalvaluation@OData.Community.Display.V1.FormattedValue" }
        ];
        createTableOfEntityLinkedToEntity("cashTable", "_js_account_value", accountId, "js_asset", cashValues);
        incomeFilter();
        assetFilter();
        getLookupFieldName("solicitor", "js_solicitor");
        getLookupFieldName("accountant", "js_accountant");

        //Facts
        factValues = [
            { width: "250px", field: "_js_contact_value@OData.Community.Display.V1.FormattedValue" },
            { width: "150px", field: "js_significance@OData.Community.Display.V1.FormattedValue" },
            { width: "450px", field: "js_comments" }
        ];
        createTableOfEntityLinkedToEntity("factTable", "_js_account_value", accountId, "js_keyfact", factValues);

        //Feedback
        feedbackValues = [
            { width: "250px", field: "js_category@OData.Community.Display.V1.FormattedValue" },
            { width: "150px", field: "js_rating@OData.Community.Display.V1.FormattedValue" },
            { width: "450px", field: "js_comments" }
        ];
        createTableOfEntityLinkedToEntity("feedbackTable", "_js_account_value", accountId, "js_feedback", feedbackValues);

        //Our Actions
        ourActionsValues = [
            { width: "600px", field: "description" },
            { width: "150px", field: "js_duedate@OData.Community.Display.V1.FormattedValue" }
        ];
        createTableOfEntityLinkedToEntity("ourActionTable", "js_type ne 100000001 and _regardingobjectid_value", accountId, "task", ourActionsValues);

        //Client Actions
        clientActionsValues = [
            { width: "600px", field: "description" },
            { width: "150px", field: "js_duedate@OData.Community.Display.V1.FormattedValue" }
        ];
        createTableOfEntityLinkedToEntity("clientActionTable", "js_type eq 100000001 and _regardingobjectid_value", accountId, "task", clientActionsValues);

        // Hide the loading icon
    setTimeout(() => {
        document.getElementById("overview").style.display = "inline";
        document.getElementById(elementId).style.display = "none";
    }, 2000); 
}


function getContact(guid) {
    return window.parent.Xrm.WebApi.retrieveRecord("contact", guid);
}

window.onload = setInitialFocus;