sap.ui.define(
  [
    "zgsp26/conf/mng/filimit/confmngfefilimit/ext/shared/ConfMngBaseController",
    "sap/m/MessageToast",
    "sap/m/MessageBox",
    "sap/ui/model/json/JSONModel",
    "sap/ui/model/Filter",
    "sap/ui/model/FilterOperator",
    "sap/m/SelectDialog",
    "sap/m/StandardListItem",
    "./ExcelImport",
  ],
  function (
    BaseController,
    MessageToast,
    MessageBox,
    JSONModel,
    Filter,
    FilterOperator,
    SelectDialog,
    StandardListItem,
    ExcelImport,
  ) {
    "use strict";

    return BaseController.extend(
      "zgsp26.conf.mng.filimit.confmngfefilimit.ext.main.Main",
      {
        // ── Abstract overrides ────────────────────────────────────────────
        _getModuleId:  function () { return "FI"; },
        _getConfName:  function () { return "FI Limit Configuration"; },
        _getTargetCds: function () { return "ZI_FI_LIMIT_CONF"; },

        // ── onInit ───────────────────────────────────────────────────────
        onInit: function () {
          this._oModel = this.getView().getModel() || this.getOwnerComponent().getModel();

          // JSONModel for table (delta approach — only changed rows go to req table)
          this._aMainRows  = [];
          this._aFieldDef  = [];   // loaded from ZCONFFIELDDEF via catalog service
          this.getView().setModel(new JSONModel({ rows: [] }), "tableData");

          const oUIModel = new JSONModel({
            editMode: true,
            requestCreated: false,
            viewOnly: false,
            totalCount: 0,
            createCount: 0,
            updateCount: 0,
            deleteCount: 0,
            rowCountText: "",
          });
          this.getView().setModel(oUIModel, "ui");

          const oFilterModel = new JSONModel({
            EnvId: "", ExpenseType: "", GlAccount: "", Currency: "", Search: "",
          });
          this.getView().setModel(oFilterModel, "filter");

          const oRequestContext = this._getUrlParams();
          this.getView().setModel(new JSONModel(oRequestContext), "requestContext");

          const bHasReqId  = !!oRequestContext.ReqId;
          const sStatus    = oRequestContext.Status || (bHasReqId ? "Draft" : "Not Created");
          const oRequestModel = new JSONModel({
            ReqId:       oRequestContext.ReqId || "",
            ReqItemId:   "",
            ConfId:      oRequestContext.ConfId || "",
            ReqTitle:    "",
            Status:      sStatus,
            StatusState: this._statusToState(sStatus),
            Reason:      "",
          });
          this.getView().setModel(oRequestModel, "request");

          if (oRequestContext.ConfId) {
            this._fetchFieldDef(oRequestContext.ConfId);
          }

          if (bHasReqId) {
            const bIsDraft    = sStatus.toUpperCase() === "DRAFT";
            const bIsRejected = sStatus.toUpperCase() === "REJECTED";
            const bCanEdit    = bIsDraft || bIsRejected;
            oUIModel.setProperty("/requestCreated", true);
            oUIModel.setProperty("/editMode", bIsDraft);   // auto-edit only for Draft
            oUIModel.setProperty("/viewOnly", !bCanEdit);  // lock Edit for Submitted/Approved/etc.
            this._fetchRequestHeader(oRequestContext.ReqId);
            this._fetchReqItem(oRequestContext.ReqId);
            this._loadMainTableWithOverlay(oRequestContext.EnvId, oRequestContext.ReqId);
          } else if (oRequestContext.Mode === "VIEW") {
            oUIModel.setProperty("/editMode", false);
            oUIModel.setProperty("/viewOnly", true);
            this._loadMainTable(oRequestContext.EnvId || "DEV");
          }

          console.log("requestContext =", oRequestContext);
        },

        // ── onCancel ─────────────────────────────────────────────────────
        onCancel: async function () {
          const oView  = this.getView();
          const sReqId = oView.getModel("request").getProperty("/ReqId");
          const sEnvId = oView.getModel("requestContext").getProperty("/EnvId") || "DEV";

          oView.setBusy(true);
          try {
            if (sReqId) {
              await this._loadMainTableWithOverlay(sEnvId, sReqId);
            } else {
              oView.getModel("tableData").setProperty("/rows", this._aMainRows.map(r => ({ ...r })));
            }
            this._applyBaseFilter();
          } finally {
            oView.setBusy(false);
          }
          oView.getModel("ui").setProperty("/editMode", false);
          MessageToast.show("Changes cancelled");
        },

        // ── onCreate ─────────────────────────────────────────────────────
        onCreate: function () {
          const sReqId = this.getView().getModel("request").getProperty("/ReqId");
          if (!sReqId) {
            MessageBox.warning("Please create a request first before adding limit lines.");
            return;
          }

          const envId = this.getView().getModel("requestContext").getProperty("/EnvId") || "DEV";
          const oTableModel = this.getView().getModel("tableData");
          const aRows       = oTableModel.getProperty("/rows");
          aRows.unshift({
            ItemId:      "",
            EnvId:       envId,
            ExpenseType: "",
            GlAccount:   "",
            AutoApprLim: 0,
            Currency:    "VND",
            ChangeNote:  "",
            VersionNo:   0,
            ActionType:  "C",
            _state:      "new",
            _reqItemId:  null,
          });
          oTableModel.setProperty("/rows", aRows);
          this._applyBaseFilter();
          MessageToast.show("New row added — fill in the fields and save.");
        },

        // ── Import from Excel ─────────────────────────────────────────────

        _xlsxPromise: null,

        _ensureXlsxLoaded: function () {
          if (window.XLSX) return Promise.resolve();
          if (this._xlsxPromise) return this._xlsxPromise;

          var sPath = sap.ui.require.toUrl(
            "zgsp26/conf/mng/filimit/confmngfefilimit/lib/xlsx.min"
          ) + ".js";

          this._xlsxPromise = new Promise(function (resolve, reject) {
            var oScript = document.createElement("script");
            oScript.src = sPath;
            oScript.onload = resolve;
            oScript.onerror = function () { reject(new Error("Failed to load SheetJS library")); };
            document.head.appendChild(oScript);
          });
          return this._xlsxPromise;
        },

        onImportFromExcel: function () {
          var that = this;
          var sReqId = this.getView().getModel("request").getProperty("/ReqId");
          if (!sReqId) {
            MessageBox.warning("Please create a request first before importing.");
            return;
          }

          this._ensureXlsxLoaded().then(function () {
            var oInput = document.createElement("input");
            oInput.type = "file";
            oInput.accept = ".xls,.xlsx,.csv";
            oInput.style.display = "none";
            document.body.appendChild(oInput);

            oInput.addEventListener("change", function (oEvent) {
              var oFile = oEvent.target.files[0];
              if (!oFile) { document.body.removeChild(oInput); return; }

              // File size check (max 5 MB)
              var MAX_FILE_SIZE = 5 * 1024 * 1024;
              if (oFile.size > MAX_FILE_SIZE) {
                MessageBox.error("File size exceeds 5 MB limit. Please use a smaller file.");
                document.body.removeChild(oInput);
                return;
              }

              var oReader = new FileReader();
              oReader.onload = function (e) {
                try {
                  that._processExcelData(e.target.result);
                } catch (err) {
                  MessageBox.error("Could not read the Excel file: " + err.message);
                }
                document.body.removeChild(oInput);
              };
              oReader.onerror = function () {
                MessageBox.error("Failed to read file.");
                document.body.removeChild(oInput);
              };
              oReader.readAsArrayBuffer(oFile);
            });

            oInput.click();
          }).catch(function (err) {
            MessageBox.error("Failed to load Excel library: " + err.message);
          });
        },

        _processExcelData: function (arrayBuffer) {
          var workbook = XLSX.read(arrayBuffer, { type: "array" });
          var sEnvId = this.getView().getModel("requestContext").getProperty("/EnvId") || "DEV";
          var result = ExcelImport.parseWorkbook(workbook, sEnvId);

          if (result.errors.length && !result.rows.length) {
            MessageBox.error(result.errors.join("\n"));
            return;
          }

          if (result.rows.length > 1000) {
            MessageBox.error("The file contains " + result.rows.length + " rows, exceeding the 1,000-row limit. Please split the file.");
            return;
          }

          if (result.rows.length > 500) {
            var that = this;
            MessageBox.confirm(
              "The file contains " + result.rows.length + " rows. This may take a moment. Continue?",
              {
                onClose: function (sAction) {
                  if (sAction === "OK") { that._insertImportedRows(result); }
                }
              }
            );
            return;
          }

          this._insertImportedRows(result);
        },

        _insertImportedRows: function (result) {
          var oTableModel = this.getView().getModel("tableData");
          var aRows = oTableModel.getProperty("/rows");

          for (var i = result.rows.length - 1; i >= 0; i--) {
            aRows.unshift(result.rows[i]);
          }

          // Validate ALL rows (imported + existing) — catches duplicates cross-set
          var aErrors = this._validateRows(aRows);
          oTableModel.setProperty("/rows", aRows);
          this._applyBaseFilter();

          var sMsg = result.rows.length + " row(s) imported from Excel.";
          if (result.skipped > 0) {
            sMsg += "\n" + result.skipped + " row(s) skipped (empty).";
          }

          if (aErrors.length) {
            MessageBox.warning(
              sMsg + "\n\n" + aErrors.length + " validation issue(s) found. Please review highlighted rows before saving.",
              { title: "Import Complete with Warnings" }
            );
          } else {
            MessageToast.show(sMsg);
          }
        },

        // ── onDelete ─────────────────────────────────────────────────────
        onDelete: function () {
          const oTable = this.byId("routesTable");
          const aItems = oTable.getSelectedItems();
          if (!aItems.length) { MessageToast.show("Select row(s) to delete"); return; }

          MessageBox.confirm("Delete selected row(s)?", {
            onClose: (sAction) => {
              if (sAction !== "OK") return;

              const oTableModel = this.getView().getModel("tableData");
              const aRows       = oTableModel.getProperty("/rows");

              aItems.forEach(oItem => {
                const oCtx = oItem.getBindingContext("tableData");
                if (!oCtx) return;
                const oRow = oCtx.getObject();
                const iIdx = aRows.indexOf(oRow);
                if (iIdx < 0) return;

                if (oRow._state === "new") {
                  aRows.splice(iIdx, 1);
                } else {
                  aRows[iIdx]._state     = "deleted";
                  aRows[iIdx].ActionType = "X";
                }
              });

              oTableModel.setProperty("/rows", aRows);
              this._applyBaseFilter();
              MessageToast.show("Marked for deletion — save to confirm.");
            },
          });
        },

        // ── onSave ───────────────────────────────────────────────────────
        onSave: async function () {
          const oView         = this.getView();
          const oRequestModel = oView.getModel("request");
          const sReqId        = oRequestModel.getProperty("/ReqId");
          const sTitle        = oRequestModel.getProperty("/ReqTitle") || "";
          const sReason       = oRequestModel.getProperty("/Reason")   || "";

          // ── Frontend validation ───────────────────────────────────────
          const oTableModelV = oView.getModel("tableData");
          const aAllRows     = oTableModelV.getProperty("/rows");
          const aErrors      = this._validateRows(aAllRows);
          oTableModelV.setProperty("/rows", aAllRows);

          if (aErrors.length > 0) {
            MessageBox.error(
              "Please fix the following errors before saving:\n\n" + aErrors.join("\n"),
              { title: "Validation Error" }
            );
            return;
          }

          oView.setBusy(true);
          try {
            const sCsrfToken = await this._fetchCsrfToken(
              this._getFiLimitServiceUrl() + "?sap-client=" + this._getSapClient()
            );
            if (!sCsrfToken) throw new Error("Cannot fetch CSRF token");

            const oTableModel = oView.getModel("tableData");
            const aRows       = oTableModel.getProperty("/rows");

            const sReqItemId = oRequestModel.getProperty("/ReqItemId") || "";
            const sConfId    = oRequestModel.getProperty("/ConfId")    || "";

            if (!sReqItemId) {
              oView.setBusy(false);
              MessageBox.error("Cannot save: Request item data could not be loaded from the server. Please refresh the page and try again.");
              return;
            }

            // Build lookup of original main table values (for OldXxx fields)
            const mOrig = {};
            this._aMainRows.forEach(r => { mOrig[r.ItemId] = r; });

            let iSaved = 0;

            for (const row of aRows) {

              // ── NEW row (ActionType = 'C') ────────────────────────────
              if (row._state === "new") {
                const oCreated = await this._postAndActivateReqRow({
                  ReqId:       sReqId,
                  ReqItemId:   sReqItemId,
                  ConfId:      sConfId,
                  ActionType:  "C",
                  EnvId:       row.EnvId       || "",
                  ExpenseType: row.ExpenseType || "",
                  GlAccount:   row.GlAccount   || "",
                  AutoApprLim: parseFloat(row.AutoApprLim) || 0,
                  Currency:    row.Currency     || "",
                  ChangeNote:  row.ChangeNote   || "",
                }, sCsrfToken);
                if (oCreated?.ItemId) {
                  row._state     = "new";
                  row._reqItemId = { ReqId: oCreated.ReqId, ReqItemId: oCreated.ReqItemId, ItemId: oCreated.ItemId };
                  iSaved++;
                }
                continue;
              }

              // ── DELETED row (ActionType = 'X') ────────────────────────
              if (row._state === "deleted") {
                if (!row._reqItemId) {
                  const orig = mOrig[row.ItemId] || row;
                  await this._postAndActivateReqRow({
                    ReqId:        sReqId,
                    ReqItemId:    sReqItemId,
                    ConfId:       sConfId,
                    ActionType:   "X",
                    SourceItemId: row.ItemId,
                    EnvId:        row.EnvId       || "",
                    ExpenseType:  row.ExpenseType || "",
                    GlAccount:    row.GlAccount   || "",
                    AutoApprLim:  parseFloat(row.AutoApprLim) || 0,
                    Currency:     row.Currency     || "",
                    VersionNo:    row.VersionNo   || 0,
                    ChangeNote:   row.ChangeNote   || "",
                    OldEnvId:        orig.EnvId       || "",
                    OldExpenseType:  orig.ExpenseType || "",
                    OldGlAccount:    orig.GlAccount   || "",
                    OldAutoApprLim:  parseFloat(orig.AutoApprLim) || 0,
                    OldCurrency:     orig.Currency     || "",
                    OldVersionNo:    orig.VersionNo   || 0,
                  }, sCsrfToken);
                  iSaved++;
                } else {
                  await this._editPatchActivateReqRow(
                    row._reqItemId,
                    { ActionType: "X" },
                    sCsrfToken
                  );
                  iSaved++;
                }
                continue;
              }

              // ── UNCHANGED row — check if user modified any field ───────
              if (!row.ItemId) continue;
              const orig = mOrig[row.ItemId];
              if (!orig) continue;

              const bChanged =
                (row.ExpenseType || "") !== (orig.ExpenseType || "") ||
                (row.GlAccount   || "") !== (orig.GlAccount   || "") ||
                (parseFloat(row.AutoApprLim) || 0) !== (parseFloat(orig.AutoApprLim) || 0) ||
                (row.Currency    || "") !== (orig.Currency    || "") ||
                !!(row.ChangeNote || "").trim();

              if (!bChanged) continue;

              if (!row._reqItemId) {
                const oCreated = await this._postAndActivateReqRow({
                  ReqId:        sReqId,
                  ReqItemId:    sReqItemId,
                  ConfId:       sConfId,
                  ActionType:   "U",
                  SourceItemId: row.ItemId,
                  EnvId:        row.EnvId       || "",
                  ExpenseType:  row.ExpenseType || "",
                  GlAccount:    row.GlAccount   || "",
                  AutoApprLim:  parseFloat(row.AutoApprLim) || 0,
                  Currency:     row.Currency     || "",
                  VersionNo:    orig.VersionNo  || 0,
                  ChangeNote:   row.ChangeNote   || "",
                  OldEnvId:        orig.EnvId       || "",
                  OldExpenseType:  orig.ExpenseType || "",
                  OldGlAccount:    orig.GlAccount   || "",
                  OldAutoApprLim:  parseFloat(orig.AutoApprLim) || 0,
                  OldCurrency:     orig.Currency     || "",
                  OldVersionNo:    orig.VersionNo   || 0,
                }, sCsrfToken);
                if (oCreated?.ItemId) {
                  row._state     = "modified";
                  row._reqItemId = { ReqId: oCreated.ReqId, ReqItemId: oCreated.ReqItemId, ItemId: oCreated.ItemId };
                  row.ActionType = "U";
                  iSaved++;
                }
              } else {
                await this._editPatchActivateReqRow(
                  row._reqItemId,
                  {
                    ExpenseType: row.ExpenseType || "",
                    GlAccount:   row.GlAccount   || "",
                    AutoApprLim: parseFloat(row.AutoApprLim) || 0,
                    Currency:    row.Currency     || "",
                    ChangeNote:  row.ChangeNote   || "",
                  },
                  sCsrfToken
                );
                iSaved++;
              }
            }

            oTableModel.setProperty("/rows", aRows);
            if (sReqId) await this._patchRequestHeader(sReqId, sTitle, sReason);

            oView.getModel("ui").setProperty("/editMode", false);
            this._applyBaseFilter();
            MessageToast.show(iSaved > 0
              ? iSaved + " change(s) saved to request."
              : "No changes detected.");
          } catch (e) {
            console.error(e);
            MessageBox.error(e?.message || "Save failed");
          } finally {
            oView.setBusy(false);
          }
        },

        // ── onRefresh ────────────────────────────────────────────────────
        onRefresh: async function () {
          const sReqId = this.getView().getModel("request").getProperty("/ReqId");
          const sEnvId = this.getView().getModel("requestContext").getProperty("/EnvId") || "DEV";
          this.getView().setBusy(true);
          try {
            if (sReqId) {
              await this._loadMainTableWithOverlay(sEnvId, sReqId);
            } else {
              await this._loadMainTable(sEnvId);
            }
            this._applyBaseFilter();
            MessageToast.show("Refreshed");
          } catch (e) {
            MessageBox.error("Refresh failed");
          } finally {
            this.getView().setBusy(false);
          }
        },

        // ── onSearch ─────────────────────────────────────────────────────
        onSearch: function () {
          const oTable      = this.byId("routesTable");
          const oBinding    = oTable.getBinding("items");
          const oFilterData = this.getView().getModel("filter").getData();

          const aFilters = [];

          if (oFilterData.EnvId)
            aFilters.push(new Filter("EnvId", FilterOperator.Contains, oFilterData.EnvId));
          if (oFilterData.ExpenseType)
            aFilters.push(new Filter("ExpenseType", FilterOperator.Contains, oFilterData.ExpenseType));
          if (oFilterData.GlAccount)
            aFilters.push(new Filter("GlAccount", FilterOperator.Contains, oFilterData.GlAccount));
          if (oFilterData.Currency)
            aFilters.push(new Filter("Currency", FilterOperator.Contains, oFilterData.Currency));

          const sSearch = this.byId("searchField").getValue();
          if (sSearch) {
            aFilters.push(new Filter({
              filters: [
                new Filter("EnvId",       FilterOperator.Contains, sSearch),
                new Filter("ExpenseType", FilterOperator.Contains, sSearch),
                new Filter("GlAccount",   FilterOperator.Contains, sSearch),
                new Filter("Currency",    FilterOperator.Contains, sSearch),
              ],
              and: false,
            }));
          }

          oBinding.filter(aFilters);
        },

        // ── onClearFilters ───────────────────────────────────────────────
        onClearFilters: function () {
          this.getView().getModel("filter").setData({
            EnvId: "", ExpenseType: "", GlAccount: "", Currency: "", Search: "",
          });
          this.byId("searchField").setValue("");
          this.byId("currencyInput").setValue("");
          this._applyBaseFilter();
          MessageToast.show("Filters cleared");
        },

        // ── Value helps (filter panel) ────────────────────────────────────

        onValueHelpEnv: function () {
          if (!this._oEnvDialog) {
            this._oEnvDialog = new SelectDialog({
              title: "Select Environment",
              liveChange: function (oEvent) {
                const sValue = oEvent.getParameter("value");
                oEvent.getSource().getBinding("items").filter([
                  new Filter("EnvId", FilterOperator.Contains, sValue),
                ]);
              },
              search: function (oEvent) {
                const sValue = oEvent.getParameter("value");
                oEvent.getSource().getBinding("items").filter([
                  new Filter("EnvId", FilterOperator.Contains, sValue),
                ]);
              },
              confirm: (oEvent) => {
                const oSelected = oEvent.getParameter("selectedItem");
                if (oSelected) {
                  this.getView().getModel("filter").setProperty("/EnvId", oSelected.getTitle());
                }
              },
            });
            this._oEnvDialog.bindAggregation("items", {
              path: "/EnvDef",
              template: new StandardListItem({ title: "{EnvId}", description: "{Description}" }),
            });
            this.getView().addDependent(this._oEnvDialog);
          }
          this._oEnvDialog.open();
        },

        onValueHelpExpenseType: function () {
          if (!this._oExpenseTypeFilterDialog) {
            this._oExpenseTypeFilterDialog = new SelectDialog({
              title: "Select Expense Type",
              liveChange: function (oEv) {
                var sVal = oEv.getParameter("value");
                oEv.getSource().getBinding("items").filter([
                  new Filter("ExpenseType", FilterOperator.Contains, sVal),
                ]);
              },
              search: function (oEv) {
                var sVal = oEv.getParameter("value");
                oEv.getSource().getBinding("items").filter([
                  new Filter("ExpenseType", FilterOperator.Contains, sVal),
                ]);
              },
              confirm: (oEv) => {
                var oItem = oEv.getParameter("selectedItem");
                if (oItem) {
                  this.getView().getModel("filter").setProperty("/ExpenseType", oItem.getTitle());
                }
              },
            });
            this._oExpenseTypeFilterDialog.bindAggregation("items", {
              path: "/ExpenseTypeDef",
              template: new StandardListItem({ title: "{ExpenseType}", description: "{Description}" }),
            });
            this.getView().addDependent(this._oExpenseTypeFilterDialog);
          }
          this._oExpenseTypeFilterDialog.open();
        },

        // ── Value helps (table cell editing) ─────────────────────────────

        _setCellValue: function (sValue) {
          if (!this._oCellCtx) return;
          this._oCellCtx.getModel().setProperty(
            this._oCellCtx.getPath() + "/" + this._sCellField,
            sValue
          );
        },

        onValueHelpExpenseTypeCell: function (oEvent) {
          this._oCellCtx   = oEvent.getSource().getBindingContext("tableData");
          this._sCellField = "ExpenseType";
          if (!this._oExpenseTypeCellDialog) {
            this._oExpenseTypeCellDialog = new SelectDialog({
              title: "Select Expense Type",
              liveChange: function (oEv) {
                var sVal = oEv.getParameter("value");
                oEv.getSource().getBinding("items").filter([
                  new Filter("ExpenseType", FilterOperator.Contains, sVal),
                ]);
              },
              search: function (oEv) {
                var sVal = oEv.getParameter("value");
                oEv.getSource().getBinding("items").filter([
                  new Filter("ExpenseType", FilterOperator.Contains, sVal),
                ]);
              },
              confirm: (oEv) => {
                var oItem = oEv.getParameter("selectedItem");
                if (oItem) this._setCellValue(oItem.getTitle());
              },
            });
            this._oExpenseTypeCellDialog.bindAggregation("items", {
              path: "/ExpenseTypeDef",
              template: new StandardListItem({ title: "{ExpenseType}", description: "{Description}" }),
            });
            this.getView().addDependent(this._oExpenseTypeCellDialog);
          }
          this._oExpenseTypeCellDialog.open();
        },

        onValueHelpCurrency: function () {
          if (!this._oCurrencyFilterDialog) {
            this._oCurrencyFilterDialog = new SelectDialog({
              title: "Select Currency",
              liveChange: function (oEv) {
                var sVal = oEv.getParameter("value");
                oEv.getSource().getBinding("items").filter([
                  new Filter("Currency", FilterOperator.Contains, sVal),
                ]);
              },
              search: function (oEv) {
                var sVal = oEv.getParameter("value");
                oEv.getSource().getBinding("items").filter([
                  new Filter("Currency", FilterOperator.Contains, sVal),
                ]);
              },
              confirm: (oEv) => {
                var oItem = oEv.getParameter("selectedItem");
                if (oItem) {
                  this.getView().getModel("filter").setProperty("/Currency", oItem.getTitle());
                }
              },
            });
            this._oCurrencyFilterDialog.bindAggregation("items", {
              path: "/CurrencyDef",
              template: new StandardListItem({ title: "{Currency}", description: "{CurrencyName}" }),
            });
            this.getView().addDependent(this._oCurrencyFilterDialog);
          }
          this._oCurrencyFilterDialog.open();
        },

        onValueHelpGlAccount: function () {
          if (!this._oGlAccountFilterDialog) {
            this._oGlAccountFilterDialog = new SelectDialog({
              title: "Select GL Account",
              liveChange: function (oEv) {
                var sVal = oEv.getParameter("value");
                oEv.getSource().getBinding("items").filter([
                  new Filter("GlAccount", FilterOperator.Contains, sVal),
                ]);
              },
              search: function (oEv) {
                var sVal = oEv.getParameter("value");
                oEv.getSource().getBinding("items").filter([
                  new Filter("GlAccount", FilterOperator.Contains, sVal),
                ]);
              },
              confirm: (oEv) => {
                var oItem = oEv.getParameter("selectedItem");
                if (oItem) {
                  this.getView().getModel("filter").setProperty("/GlAccount", oItem.getTitle());
                }
              },
            });
            this._oGlAccountFilterDialog.bindAggregation("items", {
              path: "/GlAccountDef",
              template: new StandardListItem({ title: "{GlAccount}", description: "{Description}" }),
            });
            this.getView().addDependent(this._oGlAccountFilterDialog);
          }
          this._oGlAccountFilterDialog.open();
        },

        onValueHelpCurrencyCell: function (oEvent) {
          this._oCellCtx   = oEvent.getSource().getBindingContext("tableData");
          this._sCellField = "Currency";
          if (!this._oCurrencyCellDialog) {
            this._oCurrencyCellDialog = new SelectDialog({
              title: "Select Currency",
              liveChange: function (oEv) {
                var sVal = oEv.getParameter("value");
                oEv.getSource().getBinding("items").filter([
                  new Filter("Currency", FilterOperator.Contains, sVal),
                ]);
              },
              search: function (oEv) {
                var sVal = oEv.getParameter("value");
                oEv.getSource().getBinding("items").filter([
                  new Filter("Currency", FilterOperator.Contains, sVal),
                ]);
              },
              confirm: (oEv) => {
                var oItem = oEv.getParameter("selectedItem");
                if (oItem) this._setCellValue(oItem.getTitle());
              },
            });
            this._oCurrencyCellDialog.bindAggregation("items", {
              path: "/CurrencyDef",
              template: new StandardListItem({ title: "{Currency}", description: "{CurrencyName}" }),
            });
            this.getView().addDependent(this._oCurrencyCellDialog);
          }
          this._oCurrencyCellDialog.open();
        },

        onValueHelpGlAccountCell: function (oEvent) {
          this._oCellCtx   = oEvent.getSource().getBindingContext("tableData");
          this._sCellField = "GlAccount";
          if (!this._oGlAccountCellDialog) {
            this._oGlAccountCellDialog = new SelectDialog({
              title: "Select GL Account",
              liveChange: function (oEv) {
                var sVal = oEv.getParameter("value");
                oEv.getSource().getBinding("items").filter([
                  new Filter("GlAccount", FilterOperator.Contains, sVal),
                ]);
              },
              search: function (oEv) {
                var sVal = oEv.getParameter("value");
                oEv.getSource().getBinding("items").filter([
                  new Filter("GlAccount", FilterOperator.Contains, sVal),
                ]);
              },
              confirm: (oEv) => {
                var oItem = oEv.getParameter("selectedItem");
                if (oItem) this._setCellValue(oItem.getTitle());
              },
            });
            this._oGlAccountCellDialog.bindAggregation("items", {
              path: "/GlAccountDef",
              template: new StandardListItem({ title: "{GlAccount}", description: "{Description}" }),
            });
            this.getView().addDependent(this._oGlAccountCellDialog);
          }
          this._oGlAccountCellDialog.open();
        },

        // ── Formatters ───────────────────────────────────────────────────

        formatAmount: function (val) {
          if (val === null || val === undefined || val === "") return "—";
          var num = parseFloat(val);
          if (isNaN(num)) return val;
          return num.toLocaleString("en-US", {
            minimumFractionDigits: 0,
            maximumFractionDigits: 2,
          });
        },

        // ── Statistics ───────────────────────────────────────────────────

        _updateStats: function () {
          var oUI = this.getView().getModel("ui");
          var aRows = this.getView().getModel("tableData").getProperty("/rows") || [];
          var iCreate = 0, iUpdate = 0, iDelete = 0;
          aRows.forEach(function (r) {
            if (r.ActionType === "C") iCreate++;
            else if (r.ActionType === "U") iUpdate++;
            else if (r.ActionType === "X") iDelete++;
          });
          oUI.setProperty("/totalCount", aRows.length);
          oUI.setProperty("/createCount", iCreate);
          oUI.setProperty("/updateCount", iUpdate);
          oUI.setProperty("/deleteCount", iDelete);

          var iVisible = aRows.filter(function (r) { return r._state !== "deleted"; }).length;
          oUI.setProperty("/rowCountText", iVisible + " lines");
        },

        onTableUpdateFinished: function () {
          this._updateStats();
        },

        // ── Private helpers (FI Limit specific) ──────────────────────────

        _validateRows: function (aRows) {
          const aErrors  = [];
          const aFieldDef = this._aFieldDef || [];

          const aRequired = aFieldDef.length > 0
            ? aFieldDef.filter(function (f) { return f.IsRequired; })
            : [
                { FieldName: "EXPENSE_TYPE",  FieldLabel: "Expense Type" },
                { FieldName: "GL_ACCOUNT",    FieldLabel: "GL Account" },
                { FieldName: "AUTO_APPR_LIM", FieldLabel: "Auto Approval Limit" },
              ];

          aRows.forEach(function (row, iIdx) {
            Object.keys(row).forEach(function (k) {
              if (k.startsWith("_vs_")) row[k] = "None";
            });

            const sRowNum = "Row " + (iIdx + 1);

            if (row._state !== "deleted") {
              aRequired.forEach(function (fieldDef) {
                const sJsKey = this._dbToPascal(fieldDef.FieldName);
                const sVsKey = "_vs_" + sJsKey;
                const sVal   = String(row[sJsKey] ?? "").trim();
                if (!sVal) {
                  row[sVsKey] = "Error";
                  aErrors.push(sRowNum + ": " + fieldDef.FieldLabel + " is required.");
                }
              }.bind(this));
            }
          }.bind(this));

          // ── Duplicate business key check (ExpenseType + GlAccount) ────────
          const aSeen = [];
          aRows.forEach(function (row, iIdx) {
            if (row._state === "deleted") return;
            const sExpType = String(row.ExpenseType || "").trim();
            const sGlAcc   = String(row.GlAccount   || "").trim();
            if (!sExpType || !sGlAcc) return;

            const sKey = sExpType + "|" + sGlAcc;
            if (aSeen.indexOf(sKey) >= 0) {
              row._vs_ExpenseType = "Error";
              row._vs_GlAccount   = "Error";
              aErrors.push("Row " + (iIdx + 1) + ": Duplicate entry — ExpenseType '" + sExpType + "' + GL Account '" + sGlAcc + "' already exists in this request.");
            } else {
              aSeen.push(sKey);
            }
          });

          return aErrors;
        },

        _getFiLimitServiceUrl: function () {
          return "/sap/opu/odata4/sap/zui_fi_limit_conf/srvd/sap/zsd_fi_limit_conf/0001/";
        },

        _loadMainTable: async function (sEnvId) {
          const sUrl =
            this._getFiLimitServiceUrl() +
            "FILimitMain?$filter=EnvId eq '" + sEnvId + "'&$orderby=ExpenseType,GlAccount" +
            "&sap-client=" + this._getSapClient();

          const oResp = await fetch(sUrl, {
            headers: { Accept: "application/json", "X-Requested-With": "XMLHttpRequest" },
            credentials: "include",
          });
          if (!oResp.ok) throw new Error("Cannot load FILimitMain: " + oResp.status);

          const aRaw  = (await oResp.json()).value || [];
          const aRows = aRaw.map(r => ({
            ...r,
            AutoApprLim: r.AutoApprLim ?? 0,
            ActionType: "",
            ChangeNote: "",
            _state:     "unchanged",
            _reqItemId: null,
          }));

          this._aMainRows = aRows.map(r => ({ ...r }));
          this.getView().getModel("tableData").setProperty("/rows", aRows);
          this._applyBaseFilter();
        },

        _loadMainTableWithOverlay: async function (sEnvId, sReqId) {
          const sBase      = this._getFiLimitServiceUrl();
          const sSapClient = this._getSapClient();

          const [oMainResp, oReqResp] = await Promise.all([
            fetch(sBase + "FILimitMain?$filter=EnvId eq '" + sEnvId + "'&$orderby=ExpenseType,GlAccount&sap-client=" + sSapClient, {
              headers: { Accept: "application/json", "X-Requested-With": "XMLHttpRequest" },
              credentials: "include",
            }),
            fetch(sBase + "FILimitConf?$filter=ReqId eq " + sReqId + "&$select=ItemId,ReqId,ReqItemId,SourceItemId,ActionType,EnvId,ExpenseType,GlAccount,AutoApprLim,Currency,VersionNo,ChangeNote&sap-client=" + sSapClient, {
              headers: { Accept: "application/json", "X-Requested-With": "XMLHttpRequest" },
              credentials: "include",
            }),
          ]);

          const aMain = oMainResp.ok ? ((await oMainResp.json()).value || []) : [];
          const aReq  = oReqResp.ok  ? ((await oReqResp.json()).value  || []) : [];

          const isBlankId = id => !id || id === "00000000-0000-0000-0000-000000000000";

          const mReq = {};
          aReq.forEach(r => { if (!isBlankId(r.SourceItemId)) mReq[r.SourceItemId] = r; });

          const aRows = aMain.map(mainRow => {
            const req = mReq[mainRow.ItemId];
            if (req) {
              return {
                ...req,
                _state:     req.ActionType === "X" ? "deleted" : "modified",
                _reqItemId: { ReqId: req.ReqId, ReqItemId: req.ReqItemId, ItemId: req.ItemId },
              };
            }
            return { ...mainRow, AutoApprLim: mainRow.AutoApprLim ?? 0, ActionType: "", ChangeNote: "", _state: "unchanged", _reqItemId: null };
          });

          aReq.filter(r => isBlankId(r.SourceItemId)).forEach(r => {
            aRows.push({ ...r, _state: "new", _reqItemId: { ReqId: r.ReqId, ReqItemId: r.ReqItemId, ItemId: r.ItemId } });
          });

          // Promoted deletes: row was already removed from main table (request approved + applied),
          // but must still appear in history view so users can see what was deleted.
          const aMainIds = new Set(aMain.map(r => r.ItemId));
          aReq.filter(r => r.ActionType === "X" && !isBlankId(r.SourceItemId) && !aMainIds.has(r.SourceItemId))
            .forEach(r => {
              aRows.push({ ...r, _state: "deleted", _reqItemId: { ReqId: r.ReqId, ReqItemId: r.ReqItemId, ItemId: r.ItemId } });
            });

          this._aMainRows = aMain.map(r => ({ ...r }));
          this.getView().getModel("tableData").setProperty("/rows", aRows);
        },

        _postAndActivateReqRow: async function (oPayload, sCsrfToken) {
          const sBase      = this._getFiLimitServiceUrl();
          const sSapClient = this._getSapClient();
          const sNs        = "com.sap.gateway.srvd.zsd_fi_limit_conf.v0001";
          const oHeaders   = {
            "Content-Type": "application/json",
            "X-CSRF-Token": sCsrfToken,
            "X-Requested-With": "XMLHttpRequest",
            Accept: "application/json",
          };

          const oPostResp = await fetch(sBase + "FILimitConf?sap-client=" + sSapClient, {
            method: "POST",
            headers: oHeaders,
            credentials: "include",
            body: JSON.stringify(oPayload),
          });
          if (!oPostResp.ok) {
            const oErr = await oPostResp.json().catch(() => ({}));
            throw new Error(oErr?.error?.message || "POST req row failed (" + oPostResp.status + ")");
          }
          const oDraft = await oPostResp.json();

          const sDraftKey =
            "ReqId=" + oDraft.ReqId +
            ",ReqItemId=" + oDraft.ReqItemId +
            ",ItemId=" + oDraft.ItemId;
          const sActivateUrl =
            sBase + "FILimitConf(" + sDraftKey + ",IsActiveEntity=false)/" +
            sNs + ".Activate?sap-client=" + sSapClient;
          const oActResp = await fetch(sActivateUrl, {
            method: "POST",
            headers: oHeaders,
            credentials: "include",
            body: JSON.stringify({}),
          });
          if (!oActResp.ok) {
            const oActErr = await oActResp.json().catch(() => ({}));
            throw new Error(this._extractODataError(oActErr) || "Activate failed (" + oActResp.status + ")");
          }
          return oActResp.json();
        },

        _editPatchActivateReqRow: async function (oKeys, oPayload, sCsrfToken) {
          const sBase      = this._getFiLimitServiceUrl();
          const sSapClient = this._getSapClient();
          const sNs        = "com.sap.gateway.srvd.zsd_fi_limit_conf.v0001";
          const oHeaders   = {
            "Content-Type": "application/json",
            "X-CSRF-Token": sCsrfToken,
            "X-Requested-With": "XMLHttpRequest",
            Accept: "application/json",
          };

          const sActiveKey =
            "ReqId=" + oKeys.ReqId +
            ",ReqItemId=" + oKeys.ReqItemId +
            ",ItemId=" + oKeys.ItemId;

          const sEditUrl =
            sBase + "FILimitConf(" + sActiveKey + ",IsActiveEntity=true)/" +
            sNs + ".Edit?sap-client=" + sSapClient;
          const oEditResp = await fetch(sEditUrl, {
            method: "POST",
            headers: oHeaders,
            credentials: "include",
            body: JSON.stringify({ PreserveChanges: false }),
          });
          if (!oEditResp.ok) {
            const oErr = await oEditResp.json().catch(() => ({}));
            throw new Error(oErr?.error?.message || "Edit action failed (" + oEditResp.status + ")");
          }
          const oDraft = await oEditResp.json();

          const sDraftKey =
            "ReqId=" + (oDraft.ReqId || oKeys.ReqId) +
            ",ReqItemId=" + (oDraft.ReqItemId || oKeys.ReqItemId) +
            ",ItemId=" + (oDraft.ItemId || oKeys.ItemId);

          const sPatchUrl =
            sBase + "FILimitConf(" + sDraftKey + ",IsActiveEntity=false)?sap-client=" + sSapClient;
          const oPatchResp = await fetch(sPatchUrl, {
            method: "PATCH",
            headers: oHeaders,
            credentials: "include",
            body: JSON.stringify(oPayload),
          });
          if (!oPatchResp.ok) {
            const oErr = await oPatchResp.json().catch(() => ({}));
            throw new Error(oErr?.error?.message || "PATCH draft failed (" + oPatchResp.status + ")");
          }

          const sActivateUrl =
            sBase + "FILimitConf(" + sDraftKey + ",IsActiveEntity=false)/" +
            sNs + ".Activate?sap-client=" + sSapClient;
          const oActResp = await fetch(sActivateUrl, {
            method: "POST",
            headers: oHeaders,
            credentials: "include",
            body: JSON.stringify({}),
          });
          if (!oActResp.ok) {
            const oErr = await oActResp.json().catch(() => ({}));
            throw new Error(this._extractODataError(oErr) || "Activate failed (" + oActResp.status + ")");
          }
        },

        _applyBaseFilter: function () {
          const oBinding = this.byId("routesTable").getBinding("items");
          if (oBinding) oBinding.filter([]);
          this._updateStats();
        },
      }
    );
  }
);
