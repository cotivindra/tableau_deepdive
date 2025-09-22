let app = new Vue({
  el: "#app",
  data: {
    fileName: null,
    errorMessage: null,
    sheetMap: [],
    fieldMap: [],
    worksheetSources: [],
    currentTab: 1,
    showSheets: true,
    expandedDataSources: {}, // Object to track expanded state for each data source
    worksheetList: [], // New data to hold worksheets and dashboards
    dataSourceFields: [] // New data to hold fields and dependencies

  },

  computed: {
    unusedFields() {
      const unusedFieldsMap = {};
  
      // Loop through each data source
      this.sheetMap.forEach(dataSource => {
        // Get all fields in the current data source
        const allFields = dataSource.fields;
  
        // Get all fields that are used (in worksheets, dashboards, or calculations)
        const usedFields = this.getUsedFields(dataSource);
  
        // Filter out the fields that are used, leaving only unused ones
        unusedFieldsMap[dataSource.dsName] = allFields.filter(field => !usedFields.includes(field.name));
      });
  
      return unusedFieldsMap;
    }
  },
  
  
  methods: {
    stripBrackets: function (string) {
      return (string =
        string.startsWith("[") && string.endsWith("]")
          ? string.slice(1, -1)
          : string);
    },
    processFile: function (event) {
      this.errorMessage = null;
      let workbook = event.target.files[0];
      this.fileName = workbook.name;
      let type = workbook.name.split(".").slice(-1)[0];

      if (type === "twbx") {
        let zip = new JSZip();
        zip.loadAsync(workbook).then(
          (zip) => {
            const twbName = Object.keys(zip.files).find((file) =>
              file.endsWith(".twb")
            );
            const twb = zip.files[twbName];
            twb.async("string").then((content) => {
              if (!content) return (this.errorMessage = "No twb file found!");
              this.parseXML(content);
            });
          },
          () => {
            alert("Not a valid twbx file");
          }
        );
      } else if (type === "twb") {
        let reader = new FileReader();
        reader.onload = (evt) => {
          if (!evt.target.result) return (this.errorMessage = "No file found!");
          this.parseXML(evt.target.result);
        };
        reader.readAsText(workbook);
      } else {
        this.errorMessage = "File was not a twb or twbx.";
      }
    },
    parseXML: function (text) {
      this.sheetMap = [];
      this.fieldMap = [];
      let parser = new DOMParser();
      let xml = parser.parseFromString(text, "text/xml");
      this.getSheets(xml);
      this.getFields(xml);
      this.addCalcDef(xml);
      this.getFieldsFromDatasources(xml);
      console.log("TWB content to be parsed:", xml);

      // Log after processing fields from datasources
      this.populateWorksheetsAndDashboards();  // Populate worksheets and dashboards list
      this.populateFieldsAndDependencies();    // Populate fields and dependencies list
      console.log("fieldmap = ", this.fieldMap);
      console.log("sheetmap = ", this.sheetMap);
      console.log("Datasource Fields before processing:", this.dataSourceFields);

    },

    getSheets: function (xml) {
      let sheetMap = [];
      let worksheetSources = [];
      console.log("worksheets:", xml.getElementsByTagName("worksheets"));

      if (xml.getElementsByTagName("worksheets").length > 0) {
        let worksheets = xml.getElementsByTagName("worksheets")[0].children;
        for (let worksheet of worksheets) {
          let wsName = worksheet.attributes.name.nodeValue;
          worksheetSources.push({ wsName, dataSources: [] });
          let dataSources = worksheet
            .getElementsByTagName("table")[0]
            .getElementsByTagName("view")[0]
            .getElementsByTagName("datasources")[0].children;
          for (let dataSource of dataSources) {
            let dsName = dataSource.attributes.caption
              ? dataSource.attributes.caption.nodeValue
              : dataSource.attributes.name.nodeValue;
            let dsID = dataSource.attributes.name.nodeValue;
            let foundDS = sheetMap.find((d) => d.dsName == dsName);
            if (foundDS) {
              foundDS.sheets.push({ name: wsName, type: "worksheet" });
            } else {
              sheetMap.push({
                dsName,
                sheets: [{ name: wsName, type: "worksheet" }],
              });
            }
            let foundWS = worksheetSources.find((ws) => ws.wsName === wsName);
            foundWS.dataSources.push({ name: dsName, id: dsID });
          }
        }
      }
      this.worksheetSources = worksheetSources;
      if (xml.getElementsByTagName("dashboards").length > 0) {
        let dashboards = xml.getElementsByTagName("dashboards")[0].children;
        for (let dashboard of dashboards) {
          let dbName = dashboard.attributes.name.nodeValue;
          if (dashboard.getElementsByTagName("datasources").length > 0) {
            let dataSources =
              dashboard.getElementsByTagName("datasources")[0].children;
            for (let dataSource of dataSources) {
              let dsName = dataSource.attributes.caption
                ? dataSource.attributes.caption.nodeValue
                : dataSource.attributes.name.nodeValue;
              let foundDS = sheetMap.find((d) => d.dsName == dsName);
              if (foundDS) {
                foundDS.sheets.push({ name: dbName, type: "dashboard" });
              } else {
                sheetMap.push({
                  dsName,
                  sheets: [{ name: dbName, type: "dashboard" }],
                });
              }
            }
          }
          if (dashboard.getElementsByTagName("zones").length > 0) {
            let zones = dashboard.getElementsByTagName("zone");
            for (let zone of zones) {
              if (zone.attributes.name && zone.attributes.id) {
                let wsName = zone.attributes.name.nodeValue;
                let foundWS = worksheetSources.find(
                  (ws) => ws.wsName === wsName
                );
                if (foundWS) {
                  for (let source of foundWS.dataSources) {
                    let foundDS = sheetMap.find(
                      (ds) => ds.dsName === source.name
                    );
                    if (
                      !foundDS.sheets.find((sheet) => sheet.name === dbName)
                    ) {
                      foundDS.sheets.push({ name: dbName, type: "dashboard" });
                    }
                  }
                }
              }
            }
          }
        }
      }
      this.sheetMap = sheetMap;
      if (Object.keys(sheetMap).length === 0)
        this.errorMessage = "No worksheets or dashboards found.";
    },
    getFields: function (xml) {
      let fieldMap = [];
      let calcDef = [];
      if (xml.getElementsByTagName("worksheets").length > 0) {
        let worksheets = xml.getElementsByTagName("worksheets")[0].children;
        for (let worksheet of worksheets) {
          let wsName = worksheet.attributes.name.nodeValue;
          let dsDependencies = worksheet.getElementsByTagName(
            "datasource-dependencies"
          );
          for (let dataSource of dsDependencies) {
            let dsID = dataSource.attributes.datasource.nodeValue;
            // console.log("dfid = ", dsID);
            let foundDS = this.worksheetSources
              .find((ws) => ws.wsName === wsName)
              .dataSources.find((ds) => ds.id === dsID);
            if (!foundDS) continue;
            let dsName = foundDS.name;
            if (!fieldMap.find((ds) => ds.dsName === dsName))
              fieldMap.push({ dsName, fields: [] });
            let dsFields = fieldMap.find((ds) => ds.dsName === dsName);
            let columns = dataSource.getElementsByTagName("column");
            for (let column of columns) {
              let fieldName = column.attributes.caption
                ? column.attributes.caption.nodeValue
                : column.attributes.name.nodeValue;
              fieldName = this.stripBrackets(fieldName);
              let type =
                column.getElementsByTagName("calculation").length > 0
                  ? "calculation"
                  : "datasourcefield";
              let calc =
                type === "calculation" &&
                  column.getElementsByTagName("calculation")[0].attributes.formula
                  ? column.getElementsByTagName("calculation")[0].attributes
                    .formula.nodeValue
                  : null;
              let foundField = dsFields.fields.find(
                (f) => f.name === fieldName
              );
              if (foundField) {
                foundField.worksheets.push(wsName);
              } else {
                dsFields.fields.push({
                  name: fieldName,
                  type,
                  calc,
                  worksheets: [wsName],
                });
              }
            }
          }
        }
      }
      for (let ds of fieldMap) {
        ds = ds.fields.sort((a, b) => (a.name > b.name ? 1 : -1));
      }
      this.fieldMap = fieldMap;
    },
    addCalcDef: function (xml) {
      let calcDef = [];
      let calcList = [];
      if (xml.getElementsByTagName("datasources").length > 0) {
        let datasources = xml.getElementsByTagName("datasources")[0].children;
        for (let dataSource of datasources) {
          let dsName = dataSource.attributes.caption
            ? dataSource.attributes.caption.nodeValue
            : dataSource.attributes.name.nodeValue;
          let dsID = dataSource.attributes.name.nodeValue;
          calcDef.push({ dsName, dsID, columns: [] });
          let ds = calcDef.find((ds) => ds.dsName === dsName);
          let columns = dataSource.getElementsByTagName("column");
          for (let column of columns) {
            let isCalc = column.getElementsByTagName("calculation").length > 0;
            if (isCalc) {
              let name = column.attributes.caption
                ? column.attributes.caption.nodeValue
                : column.attributes.name.nodeValue;
              let id = column.attributes.name.nodeValue;
              ds.columns.push({ name, id });
            }
          }
        }
      }

      for (let ds of calcDef) {
        let dsName = ds.dsName;
        for (let calc of ds.columns) {
          let name = this.stripBrackets(calc.name);
          let id = calc.id;
          let displayDS = this.fieldMap.find((ds) => ds.dsName === dsName);
          if (!displayDS) continue;
          for (let field of displayDS.fields) {
            if (field.calc) {
              field.calc = field.calc.replaceAll(id, `[${name}]`);
              for (let ds2 of calcDef) {
                let ds2Name = this.stripBrackets(ds2.dsName);
                let ds2ID = ds2.dsID;
                if (ds2ID !== "Parameters")
                  field.calc = field.calc.replaceAll(ds2ID, `[${ds2Name}]`);
              }
              if (
                !calcList.find(
                  (f) => f.dsName === dsName && f.name === field.name
                )
              )
                calcList.push({ dsName, name: field.name, calc: field.calc });
            }
          }
        }
      }

      for (let calc of calcList) {
        if (calc.dsName !== "Parameters") {
          let name = calc.name;
          let r = new RegExp(/\[([^\[\]]+)\]/g);
          let matches = calc.calc.match(r);
          if (matches && matches.length > 0) {
            for (let match of matches) {
              let depName = this.stripBrackets(match);
              for (let ds of this.fieldMap) {
                for (let field of ds.fields) {
                  if (
                    depName === field.name &&
                    field.worksheets.indexOf(`=${name}`) === -1
                  )
                    field.worksheets.push(`=${name}`);
                }
              }
            }
          }
        }
      }
    },

    getSheetDataSourceMapping: function () {
      let mapping = {};

      for (let sheet of this.worksheetSources) {
        for (let dataSource of sheet.dataSources) {
          let wsName = sheet.wsName;
          let dsName = dataSource.name;

          if (mapping.hasOwnProperty(wsName)) {
            mapping[wsName].push(dsName);
          } else {
            mapping[wsName] = [dsName];
          }
        }
      }

      return mapping;
    },

    getDataSourcFieldNames: function () {
      let mapping = {};

      for (let dataSource of this.fieldMap) {
        let dsName = dataSource.dsName;

        //below line will show original input datafields
        let originalFields = dataSource.fields.filter(field => field.type === 'datasourcefield').map(field => field.name);
        let calculatedFields = dataSource.fields.filter(field => field.type === 'calculation').map(field => field);
        // let allFields = dataSource.fields.map(field => field.name);

        mapping[dsName] = {
          original: originalFields,
          calculated: calculatedFields
        };
      }

      return mapping;
    },
    getFieldsFromDatasources: function (xml) {
      let datasourceFields = [];  // New separate list for fields
    
      if (xml.getElementsByTagName("datasources").length > 0) {
        let datasources = xml.getElementsByTagName("datasources")[0].children;
        
        // Loop through each datasource in the XML
        for (let dataSource of datasources) {
          let dsName = dataSource.attributes.caption
            ? dataSource.attributes.caption.nodeValue
            : dataSource.attributes.name.nodeValue;
          let dsID = dataSource.attributes.name.nodeValue;
          
          // Initialize datasource structure for this datasource
          let dsFields = { dsName, dsID, columns: [] };
          
          // Loop through each column in the datasource
          let columns = dataSource.getElementsByTagName("column");
          for (let column of columns) {
            let isCalc = column.getElementsByTagName("calculation").length > 0;
            let fieldName = column.attributes.caption
              ? column.attributes.caption.nodeValue
              : column.attributes.name.nodeValue;
            fieldName = this.stripBrackets(fieldName);
            
            if (isCalc) {
              // Handle calculated field
              let calcFormula = column.getElementsByTagName("calculation")[0].attributes.formula.nodeValue;
              dsFields.columns.push({
                name: fieldName,
                id: column.attributes.name.nodeValue,
                type: "calculation",
                formula: calcFormula
              });
              
            } else {
              // Handle regular datasource field
              dsFields.columns.push({
                name: fieldName,
                id: column.attributes.name.nodeValue,
                type: "datasourcefield",
                formula: null
              });
            }
          }
    
          // Add this datasource's fields to the list
          datasourceFields.push(dsFields);
        }
      }
      // Log after processing fields from datasources
  console.log("Datasource Fields after processing:", datasourceFields);

      // Process the calculated fields to detect dependencies
      this.processCalculatedFields(datasourceFields);
    
      // Store the final field map using datasourceFields
      for (let ds of datasourceFields) {
        let displayDS = this.fieldMap.find((dsObj) => dsObj.dsName === ds.dsName);
        if (!displayDS) {
          this.fieldMap.push({ dsName: ds.dsName, fields: ds.columns });
        }
      }

       // Log after updating field map
  console.log("Field Map after adding datasource fields:", this.fieldMap);

    },
    
    processCalculatedFields: function (datasourceFields) {
      let calcList = [];
    
      // Loop through each datasource's calculated columns
      for (let ds of datasourceFields) {
        let dsName = ds.dsName;
        for (let calc of ds.columns) {
          if (calc.type === "calculation") {
            // Extract dependencies from calculation formula
            let dependencies = this.getDependenciesForCalc(calc);
            calc.dependencies = dependencies;
    
            // For each dependent field, update its calculation references
            for (let dep of dependencies) {
              // For each datasource and field, check if it depends on any of the calculated fields
              let depDS = this.fieldMap.find((ds) => ds.dsName === dep);
              if (depDS) {
                for (let field of depDS.fields) {
                  if (field.name === dep) {
                    field.worksheets = field.worksheets || [];
                    field.worksheets.push(`=${calc.name}`);
                  }
                }
              }
            }
    
            // Add to the calculation list for reference
            calcList.push({ dsName, name: calc.name, calc: calc.formula });
          }
        }
      }
      
      // Loop through the calcList to link calculations
      for (let calc of calcList) {
        let ds = this.fieldMap.find((dsObj) => dsObj.dsName === calc.dsName);
        if (!ds) continue;
    
        for (let field of ds.fields) {
          if (field.calc) {
            let calcFormula = field.calc;
            // Replace dependencies in the calculation formula
            for (let dep of calcList) {
              if (calc.dsName !== dep.dsName) {
                let depField = this.stripBrackets(dep.name);
                calcFormula = calcFormula.replace(new RegExp(`\\[${dep.name}\\]`, 'g'), `[${depField}]`);
              }
            }
            field.calc = calcFormula;
          }
        }
      }
      // Log after updating fields with linked calculations
  console.log("Field Map after linking calculations:", this.fieldMap);

    },
    
    // Helper function to extract dependencies from calculation formulas
    getDependenciesForCalc: function (calc) {
      let formula = calc.formula;
      let dependencies = [];
    
      // Use a regular expression to find all field names in the formula
      let matches = formula.match(/\[([^\]]+)\]/g);
      if (matches) {
        dependencies = matches.map((match) => this.stripBrackets(match));
      }
    
      return dependencies;
    },
    
    
    // This method populates the list of worksheets and dashboards for Tab 6
    populateWorksheetsAndDashboards: function () {
      this.worksheetList = [];
      // Loop through the sheetMap to get worksheets and dashboards
      for (let sheet of this.sheetMap) {
        for (let entry of sheet.sheets) {
          this.worksheetList.push({
            name: entry.name,
            type: entry.type // Type can be 'worksheet' or 'dashboard'
          });
        }
      }
    },


    // This method populates the fields and dependencies for Tab 7
    populateFieldsAndDependencies: function () {
      this.dataSourceFields = [];
      // Loop through each datasource in fieldMap
      for (let ds of this.fieldMap) {
        let fieldsData = { dsName: ds.dsName, original: [], calculated: [] };
        for (let field of ds.fields) {
          if (field.type === 'datasourcefield') {
            fieldsData.original.push(field.name);
          } else if (field.type === 'calculation') {
            // Get dependencies for calculated fields
            let calcDependencies = this.getDependenciesForCalc(field);
            fieldsData.calculated.push({
              name: field.name,
              dependencies: calcDependencies
            });
          }
        }
        this.dataSourceFields.push(fieldsData);
      }
    },
    
    // This helper method extracts field dependencies (fields used inside calculations)
    getDependenciesForCalc: function (field) {
      let dependencies = [];
      if (field.calc) {
        // Regex to find all fields used within the calculation
        let regex = /\[([^\[\]]+)\]/g;
        let matches;
        while ((matches = regex.exec(field.calc)) !== null) {
          dependencies.push(matches[1]);
        }
      }
      return dependencies;
    },

  // Function to get independent calculated fields
  getIndependentCalculations(calculatedFields) {
    if (!calculatedFields || !Array.isArray(calculatedFields)) return [];
    
    return calculatedFields.filter(field => {
      // Check if dependencies exists and is an array
      return (!field.dependencies || field.dependencies.length === 0);
    });
  }
  ,

  getDependentCalculations(calculatedFields) {
    if (!calculatedFields || !Array.isArray(calculatedFields)) return []; // Ensure calculatedFields is an array
    return calculatedFields.filter(field => {
      // Check if dependencies is defined and is an array with a length > 0
      return Array.isArray(field.dependencies) && field.dependencies.length > 0;
    });
  },
  
  // Function to get unique dependencies for a given calculated field
  getUniqueDependencies(dependencies) {
    return [...new Set(dependencies)];
  },

  getUniqueList(type) {
    const uniqueList = [];
    const seenNames = new Set();
    
    this.worksheetList.forEach(ws => {
      if (ws.type === type && !seenNames.has(ws.name)) {
        uniqueList.push(ws);
        seenNames.add(ws.name);
      }
    });

    return uniqueList;
  },
  
  // getUnusedFields: function () {
    // getUsedFields(dataSource) {
      // Collect all the used field names
      // let usedFields = new Set();

      // // Loop through worksheets and dashboards to collect all fields that are in use
      // this.worksheetList.forEach(ws => {
    //   // For each worksheet and dashboard, get the data source mappings
    //   let dataSources = this.getSheetDataSourceMapping()[ws.name];
    //   if (dataSources) {
      //     dataSources.forEach(dsName => {
        //       // For each data source, collect the fields used
        //       let dataSourceFields = this.getDataSourcFieldNames()[dsName];
        //       if (dataSourceFields) {
          //         dataSourceFields.original.forEach(fieldName => usedFields.add(fieldName));
          //         dataSourceFields.calculated.forEach(field => usedFields.add(field.name));
          //       }
          //     });
          //   }
          // });

    // // Now we go through fieldMap and find any field that is not used
    // let unusedFields = [];

    // this.fieldMap.forEach(ds => {
    //   ds.fields.forEach(field => {
    //     // If the field is not in the usedFields set, it is unused
    //     if (!usedFields.has(field.name)) {
    //       unusedFields.push({ dsName: ds.dsName, field: field.name });
    //     }
    //   });
    // });
    
    // return unusedFields; // Return the list of unused fields
    
    
    // let usedFields = [];
    
    // //Check fields used in worksheets
    // dataSource.sheets.forEach(sheet => {
    //   if (sheet.type === 'worksheet' || sheet.type === 'dashboard') {
    //     sheet.fields.forEach(field => {
    //       usedFields.push(field.name);
    //     });
    //   }
    // });
    
    // // Check fields used in calculated fields
    // dataSource.fields.forEach(field => {
    //   if (field.type === 'calculation') {
    //     usedFields.push(field.name);
  //     }
  //   });
    
  //   return usedFields;
    
  // },
  
    
  toggleDataExpansion: function (data) {
    this.$set(this.expandedData, data, !this.expandedData[data]);
  },
  isDataExpanded: function (data) {
    return this.expandedData[data];
  },
  
   copyToClipboard(text) {
    
    navigator.clipboard.writeText(text)
      .then(() => {
        
        const alertBox = document.createElement('div');
        alertBox.textContent = "Formula Copied!!";
        alertBox.style.position = 'fixed';
        alertBox.style.top = '20px';
        alertBox.style.left = '50%';
        alertBox.style.transform = 'translateX(-50%)';
        alertBox.style.padding = '10px 20px';
        alertBox.style.backgroundColor = '#4CAF50';
        alertBox.style.color = 'white';
        alertBox.style.borderRadius = '5px';
        alertBox.style.fontSize = '16px';
        alertBox.style.zIndex = '1000'; 
        
        
        document.body.appendChild(alertBox);
  
        
        setTimeout(() => {
          alertBox.remove();
        }, 1000);
      })
      .catch(err => {
        
        console.error('Failed to copy text: ', err);
        const alertBox = document.createElement('div');
        alertBox.textContent = "Failed to copy formula.";
        alertBox.style.position = 'fixed';
        alertBox.style.top = '20px';
        alertBox.style.left = '50%';
        alertBox.style.transform = 'translateX(-50%)';
        alertBox.style.padding = '10px 20px';
        alertBox.style.backgroundColor = '#f44336';
        alertBox.style.color = 'white';
        alertBox.style.borderRadius = '5px';
        alertBox.style.fontSize = '16px';
        alertBox.style.zIndex = '1000'; 
        
        document.body.appendChild(alertBox);
        
        setTimeout(() => {
          alertBox.remove();
        }, 1000);
      });
  },
  
  // exportToExcel() {
  //   const wb = XLSX.utils.book_new(); // Create a new workbook
  
  //   // Loop over each data source in fieldMap
  //   fieldMap.forEach(ds => {
  //     const sheetData = [];
  
  //     // Add the header row for each sheet
  //     sheetData.push(['Field Name', 'Type', 'Formula']);
  
  //     // Loop through the fields in each data source
  //     ds.fields.forEach(field => {
  //       if (field.type === 'calculation') {
  //         // Add the field details for calculated fields
  //         sheetData.push([field.name, 'Calculation', field.calc || '']);
  //       }
  //     });
  
  //     // Create a sheet for the current data source
  //     const ws = XLSX.utils.aoa_to_sheet(sheetData);
  //     XLSX.utils.book_append_sheet(wb, ws, ds.dsName);
  //   });
  
  //   // Generate and download the Excel file
  //   XLSX.writeFile(wb, 'CalculatedFields.xlsx');
  // }
  exportToExcel() {
    const wb = XLSX.utils.book_new(); // Create a new workbook
  
    // Loop through each data source in the fieldMap
    this.fieldMap.forEach(ds => {
      const sheetData = [];
  
      // Add the header row for each sheet
      sheetData.push(['Field Name', 'Type', 'Formula']);
  
      // Get the independent calculations and add them to the sheet
      const independentFields = this.getIndependentCalculations(ds.fields);
      independentFields.forEach(field => {
        sheetData.push([field.name, field.type, field.calc || '']);
      });
  
      // Get the dependent calculations and add them to the sheet
      const dependentFields = this.getDependentCalculations(ds.fields);
      dependentFields.forEach(field => {
        sheetData.push([field.name, field.type, field.calc || '']);
      });
  
      // Create a sheet for the current data source
      const ws = XLSX.utils.aoa_to_sheet(sheetData);
  
      // Apply wrap text to the "Formula" column (index 2)
      sheetData.forEach((row, rowIndex) => {
        if (rowIndex === 0) return; // Skip header row
        ws['C' + (rowIndex + 1)].s = { alignment: { wrapText: true } }; // Wrap text for "Formula" column
      });
  
      // Apply table style (Optional: You can customize this to your needs)
      const range = XLSX.utils.decode_range(ws['!ref']);
      ws['!cols'] = this.autoFitColumns(sheetData); // Auto adjust columns
  
      // Auto resize rows
      ws['!rows'] = this.autoFitRows(sheetData);
  
      // Append the sheet to the workbook
      XLSX.utils.book_append_sheet(wb, ws, ds.dsName);
    });
  
    // Generate and download the Excel file
    XLSX.writeFile(wb, 'CalculatedFields.xlsx');
  },
  
  // Auto adjust columns based on max width
  autoFitColumns(sheetData) {
    const colWidths = [];
    sheetData.forEach(row => {
      row.forEach((cell, colIndex) => {
        const cellLength = String(cell).length;
        if (!colWidths[colIndex] || cellLength > colWidths[colIndex]) {
          colWidths[colIndex] = cellLength;
        }
      });
    });
    return colWidths.map(width => ({ wch: width + 2 })); // Add padding to the column width
  },
  
  // Auto adjust rows based on wrapped text
  autoFitRows(sheetData) {
    const rowHeights = [];
    sheetData.forEach((row, rowIndex) => {
      let maxHeight = 0;
      row.forEach(cell => {
        const cellLength = String(cell).length;
        const estimatedHeight = Math.ceil(cellLength / 50); // Estimate height based on characters per line
        maxHeight = Math.max(maxHeight, estimatedHeight);
      });
      if (maxHeight > 0) rowHeights.push({ hpx: maxHeight * 15 }); // Approx 15 px per line
    });
    return rowHeights;
  },
},
});
