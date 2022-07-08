        #region GenerateExcel
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult GenerateExcelPrenomina(Parameters Parameters)
        {
            try
            {
                IRestResponse response;
                //Se generan las variables de session
                if (!SessionApi())
                {
                    hc.GuardarLogs("Variables de session NO generadas desde el metodo GenerateExcelPrenomina en el controlador Transactions.",
                                   "Proceso incompleto, no se pudieron generar las variables de sesión para el proceso",
                                   User.Identity.Name,
                                   "N/A");
                    return View("Error");
                }
                //response = new ApiUrl().UrlExecute("PaysheetNovelty/PayTimeAgreement?Agre=" + RuteCreate.Agre, Session["ApiToken"].ToString(), "POST", null);
                //int MultiplicadorMensual = Convert.ToInt32(JsonConvert.DeserializeObject<string>(response.Content));
                //if (MultiplicadorMensual == 2)
                //{
                //    MultiplicadorMensual = 2;
                //}
                //else
                //{
                //    MultiplicadorMensual = 1;
                //}
                //Aplico licencia 
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                //Inicio el documento de excel
                ExcelPackage Excel = new ExcelPackage();
                Excel.Workbook.Properties.Title = "ME plantilla nómina";
                Excel.Workbook.Properties.Author = "Misión Empresarial";
                //Creo las hojas del archivo
                var Novedades = Excel.Workbook.Worksheets.Add("Novedades");
                var Ausentismos = Excel.Workbook.Worksheets.Add("Ausentismos");
                var Retiros = Excel.Workbook.Worksheets.Add("Retiros");
                var Formulas = Excel.Workbook.Worksheets.Add("Formulas");
                Formulas.Name = "Formulas";
                //Busco la información para generar el archivo
                searchData sd = new searchData();
                //sd.Date = Convert.ToDateTime(RuteCreate.Date).ToString("yyyyMMdd");
                //sd.Agreement = RuteCreate.Agre;
                ////Estraigo los empleados del convenio de la compañia
                response = new ApiUrl().UrlExecute("PaysheetNovelty/EmployeeCostStructure", Session["ApiToken"].ToString(), "POST", JsonConvert.SerializeObject(sd));
                if (response.StatusDescription != "OK")
                {
                    hc.GuardarLogs("Error al generar el listado de conceptos para la generación del excel.", response.StatusDescription, NameCompany(), Session["NIT"].ToString());
                    Alerts("No pudimos generar el listado de empleados, inténtalo de nuevo.", NotificationTypes.error);
                    //return RedirectToAction("NoveltiesAndAbsences", new ToPrincipalView { date = RuteCreate.Date, Agree = RuteCreate.Agre, IncomeType = RuteCreate.IncomeType });
                }
                var ListEmps = JsonConvert.DeserializeObject<List<employeeCostStructure>>(response.Content);
                var ListEmp = ListEmps.OrderBy(x => x.strNombreEmpleado);
                //Extraigo los conceptos de novedad
                response = new ApiUrl().UrlExecute("Paysheet/ConceptsByCustomers?nit=" + Session["NIT"].ToString(), Session["ApiToken"].ToString(), "POST", null);
                if (response.StatusDescription != "OK")
                {
                    hc.GuardarLogs("Error al generar el listado de ausentismos para la generación del excel.", response.StatusDescription, NameCompany(), Session["NIT"].ToString());
                    Alerts("No pudimos generar el listado de conceptos, inténtalo de nuevo.", NotificationTypes.error);
                    //return RedirectToAction("NoveltiesAndAbsences", new ToPrincipalView { date = RuteCreate.Date, Agree = RuteCreate.Agre, IncomeType = RuteCreate.IncomeType });
                }
                var ListCon = JsonConvert.DeserializeObject<List<PaysheetConceptCustomer>>(response.Content);
                //Extraigo todos los tipo de ausentismos disponibles
                response = new ApiUrl().UrlExecute("Paysheet/Absense", Session["ApiToken"].ToString(), "GET", null);
                if (response.StatusDescription != "OK")
                {
                    hc.GuardarLogs("Error al generar el listado de ausentismos para la generación del excel",
                        response.StatusDescription, NameCompany(), Session["NIT"].ToString());
                    Alerts("No pudimos generar el listado de ausentismos, inténtalo de nuevo.", NotificationTypes.error);
                    //return RedirectToAction("NoveltiesAndAbsences", new ToPrincipalView { date = RuteCreate.Date, Agree = RuteCreate.Agre, IncomeType = RuteCreate.IncomeType });
                }
                var ListAbs = JsonConvert.DeserializeObject<List<AbsenseType>>(response.Content);
                //Extraigo todas las causas de retiro dispinibles 
                response = new ApiUrl().UrlExecute("Paysheet/CauseRetreat", Session["ApiToken"].ToString(), "GET", null);
                if (response.StatusDescription != "OK")
                {
                    hc.GuardarLogs("Error al generar el listado de causas de retiro para la generación del excel",
                        response.StatusDescription, NameCompany(), Session["NIT"].ToString());
                    Alerts("No pudimos generar el listado de causas de retiro, inténtalo de nuevo.", NotificationTypes.error);
                    //return RedirectToAction("NoveltiesAndAbsences", new ToPrincipalView { date = RuteCreate.Date, Agree = RuteCreate.Agre, IncomeType = RuteCreate.IncomeType });
                }
                var ListCauseRetreat = JsonConvert.DeserializeObject<List<CauseRetreat>>(response.Content);
                QueryByCustomer ForDatesPaysheet = new QueryByCustomer();
                ResponseForDatesPaysheet ResultsDatesPaysheet = new ResponseForDatesPaysheet();
                //ForDatesPaysheet.Agreement = RuteCreate.Agre.Trim();
                //ForDatesPaysheet.Date = RuteCreate.Date.Trim();
                //ForDatesPaysheet.Nit = Session["NIT"].ToString();
                response = new ApiUrl().UrlExecute("PaysheetNovelty/DateCutForCustomerAgreement", Session["ApiToken"].ToString(), "POST", JsonConvert.SerializeObject(ForDatesPaysheet));
                if (response.StatusDescription == "OK")
                {
                    ResultsDatesPaysheet = JsonConvert.DeserializeObject<ResponseForDatesPaysheet>(response.Content);
                }
                else
                {
                    hc.GuardarLogs("Error al generar las fechas finales e iniciales de la nómina para la generación del excel.", response.StatusDescription, NameCompany(), Session["NIT"].ToString());
                    Alerts("No pudimos generar el listado de conceptos, inténtalo de nuevo.", NotificationTypes.error);
                    //return RedirectToAction("NoveltiesAndAbsences", new ToPrincipalView { date = RuteCreate.Date, Agree = RuteCreate.Agre, IncomeType = RuteCreate.IncomeType });

                }
                //Comienzo llenado la lista de novedades y retiros.
                //Se deben bloquear las cedulas y nombres de las personas
                Novedades.Protection.IsProtected = true;
                int filas = 2;
                int FilasRetiros = 2;
                int Columnas = 14;
                Novedades.Cells[1, 1].Value = "CÉDULA";
                Novedades.Cells[1, 1].Style.Font.Size = 13;
                Novedades.Cells[1, 2].Value = "NOMBRE EMPLEADOS";
                Novedades.Cells[1, 2].Style.Font.Size = 13;
                Novedades.Cells[1, 3].Value = "CONVENIO";
                Novedades.Cells[1, 3].Style.Font.Size = 13;
                Novedades.Cells[1, 4].Value = "SUCURSAL";
                Novedades.Cells[1, 4].Style.Font.Size = 13;
                Novedades.Cells[1, 5].Value = "CENTRO DE COSTOS";
                Novedades.Cells[1, 5].Style.Font.Size = 13;
                Novedades.Cells[1, 6].Value = "CLASIFICADOR 1";
                Novedades.Cells[1, 6].Style.Font.Size = 13;
                Novedades.Cells[1, 7].Value = "CLASIFICADOR 2";
                Novedades.Cells[1, 7].Style.Font.Size = 13;
                Novedades.Cells[1, 8].Value = "CLASIFICADOR 3";
                Novedades.Cells[1, 8].Style.Font.Size = 13;
                Novedades.Cells[1, 9].Value = "CLASIFICADOR 4";
                Novedades.Cells[1, 9].Style.Font.Size = 13;
                Novedades.Cells[1, 10].Value = "CLASIFICADOR 5";
                Novedades.Cells[1, 10].Style.Font.Size = 13;
                Novedades.Cells[1, 11].Value = "CLASIFICADOR 6";
                Novedades.Cells[1, 11].Style.Font.Size = 13;
                Novedades.Cells[1, 12].Value = "CLASIFICADOR 7";
                Novedades.Cells[1, 12].Style.Font.Size = 13;
                Novedades.Cells[1, 13].Value = "CLASIFICADOR 8";
                Novedades.Cells[1, 13].Style.Font.Size = 13;
                foreach (var item in ListEmp)
                {
                    Novedades.Cells[filas, 1].Value = item.strCodigoEmpleado;
                    Novedades.Cells[filas, 2].Value = item.strNombreEmpleado;
                    Novedades.Cells[filas, 3].Value = item.strConvenio;
                    Novedades.Cells[filas, 4].Value = item.strSucursal;
                    Novedades.Cells[filas, 5].Value = item.strCentroDeCostos;
                    Novedades.Cells[filas, 6].Value = item.strClasificador1;
                    Novedades.Cells[filas, 7].Value = item.strClasificador2;
                    Novedades.Cells[filas, 8].Value = item.strClasificador3;
                    Novedades.Cells[filas, 9].Value = item.strClasificador4;
                    Novedades.Cells[filas, 10].Value = item.strClasificador5;
                    Novedades.Cells[filas, 11].Value = item.strClasificador6;
                    Novedades.Cells[filas, 12].Value = item.strClasificador7;
                    Novedades.Cells[filas, 13].Value = item.strClasificador8;
                    Retiros.Cells[filas, 1].Value = item.strCodigoEmpleado;
                    Retiros.Cells[filas, 2].Value = item.strNombreEmpleado;
                    filas++;
                    FilasRetiros++;
                }

                double MaxValueValidation = 0.01;
                foreach (var item in ListCon)
                {
                    Novedades.Column(Columnas).Style.Numberformat.Format = "Texto";
                    Novedades.Cells[1, Columnas].Value = item.Concept_Id + " - " + item.Description + " (" + item.Naturaleza + ")";
                    Novedades.Cells[1, Columnas].Style.Font.Size = 13;
                    if (item.Limit == "C")
                    {
                        QuerySunday Datas = new QuerySunday();
                        Datas.DateStart = Convert.ToDateTime(ResultsDatesPaysheet.DateStart).ToString("yyyyMMdd");
                        Datas.DateEnd = Convert.ToDateTime(ResultsDatesPaysheet.DateEnd).ToString("yyyyMMdd");
                        Datas.Type = "1";

                        IRestResponse responseCalculate = new ApiUrl().UrlExecute("Paysheet/GetSundays", Session["ApiToken"].ToString(), "POST", JsonConvert.SerializeObject(Datas));
                        if (responseCalculate.StatusDescription == "OK")
                        {
                            //MaxValueValidation = Convert.ToDouble(JsonConvert.DeserializeObject<string>(responseCalculate.Content)) * MultiplicadorMensual;
                        }
                        else
                        {
                            MaxValueValidation = 0.01;
                        }
                        Novedades.Cells[2, Columnas, filas - 1, Columnas].Style.Numberformat.Format = "0.00";
                        var validation = Novedades.Cells[2, Columnas, filas - 1, Columnas].DataValidation.AddDecimalDataValidation();
                        validation.ErrorStyle = ExcelDataValidationWarningStyle.stop;
                        validation.PromptTitle = "Registro de novedad";
                        validation.ErrorTitle = "El valor ingresado no es válido";
                        validation.Operator = ExcelDataValidationOperator.between;
                        validation.Formula.Value = 0.01;
                        validation.Prompt = "El valor máximo disponible es de 0,01 a " + MaxValueValidation + " Nota: Si no desea ingresar algún valor, use la tecla SUPR";
                        validation.Error = "Solo puede ingresar valores entre 0,01 y " + MaxValueValidation;
                        validation.Formula2.Value = MaxValueValidation;
                        validation.ShowInputMessage = true;
                        validation.ShowErrorMessage = true;

                    }
                    else
                    {
                        item.Limit = (Convert.ToDouble(item.Limit)).ToString();
                        Novedades.Cells[2, Columnas, filas - 1, Columnas].Style.Numberformat.Format = "0.00";
                        var validation = Novedades.Cells[2, Columnas, filas - 1, Columnas].DataValidation.AddDecimalDataValidation();
                        validation.ErrorStyle = ExcelDataValidationWarningStyle.stop;
                        validation.PromptTitle = "Registro de novedad";
                        validation.ErrorTitle = "El valor ingresado no es válido";
                        validation.Operator = ExcelDataValidationOperator.between;
                        validation.Formula.Value = 0.01;
                        validation.Prompt = "El valor máximo disponible es de 0,01 a " + item.Limit + " Nota: Si no desea ingresar algún valor use la tecla SUPR";
                        validation.Error = "Solo puede ingresar valores entre 0,01 y " + item.Limit;
                        validation.Formula2.Value = Convert.ToDouble(item.Limit);
                        validation.ShowInputMessage = true;
                        validation.ShowErrorMessage = true;
                    }

                    Columnas++;
                }
                Novedades.Cells[2, 14, filas - 1, Columnas - 1].Style.Locked = false;
                Novedades.Row(1).Height = 41.25;
                Novedades.Cells[1, 1, 1, (ListCon.Count + 13)].Style.Fill.PatternType = ExcelFillStyle.Solid;
                Novedades.Cells[1, 1, 1, (ListCon.Count + 13)].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F4F6F8"));
                Novedades.Cells[1, 1, 1, (ListCon.Count + 13)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                Novedades.Cells[1, 1, 1, (ListCon.Count + 13)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                Novedades.Cells[1, 1, 1, (ListCon.Count + 13)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                Novedades.Cells[1, 1, 1, (ListCon.Count + 13)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                Novedades.Cells[1, 1, 1, (ListCon.Count + 13)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Novedades.Cells[1, 1, 1, (ListCon.Count + 13)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                Novedades.Cells[1, 1, 1, (ListCon.Count + 13)].Style.Font.Bold = true;
                Novedades.Cells[Novedades.Dimension.Address].Style.Font.Name = "Century Gothic";
                Novedades.Cells[Novedades.Dimension.Address].AutoFitColumns();
                Novedades.View.FreezePanes(2, 3);
                //Hoja de Ausentismos
                Ausentismos.Protection.IsProtected = true;
                Ausentismos.Cells[1, 1].Value = "EMPLEADO";
                Ausentismos.Cells[1, 1].Style.Font.Size = 13;
                Ausentismos.Cells[1, 2].Value = "TIPO AUSENTISMO";
                Ausentismos.Cells[1, 2].Style.Font.Size = 13;
                Ausentismos.Cells[1, 3].Value = "FECHA INICIO";
                Ausentismos.Cells[1, 3].Style.Font.Size = 13;
                Ausentismos.Cells[1, 4].Value = "FECHA FINAL";
                Ausentismos.Cells[1, 4].Style.Font.Size = 13;
                Ausentismos.Cells[1, 5].Value = "FECHA REAL INICIO";
                Ausentismos.Cells[1, 5].Style.Font.Size = 13;
                Ausentismos.Cells[1, 6].Value = "FECHA REAL FINALIZACIÓN";
                Ausentismos.Cells[1, 6].Style.Font.Size = 13;
                filas = 2;
                foreach (var item in ListEmp)
                {
                    Formulas.Cells[filas, 1].Value = item.strCodigoEmpleado + " - " + item.strNombreEmpleado;
                    filas++;
                }
                filas = 2;
                foreach (var item in ListAbs)
                {
                    Formulas.Cells[filas, 2].Value = item.strTipoAusencia + " - " + item.strNombreAusencia;
                    filas++;
                }
                var dd = Ausentismos.Cells["A2:A500"].DataValidation.AddListDataValidation() as ExcelDataValidationList;
                dd.AllowBlank = true;
                dd.Formula.ExcelFormula = "Formulas!A$2:A" + (ListEmps.Count + 1);
                var dd2 = Ausentismos.Cells["B2:B500"].DataValidation.AddListDataValidation() as ExcelDataValidationList;
                dd2.AllowBlank = true;
                dd2.Formula.ExcelFormula = "Formulas!B$2:B$" + (ListAbs.Count + 1);
                Ausentismos.Row(1).Height = 41.25;
                Ausentismos.Cells[1, 1, 1, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                Ausentismos.Cells[1, 1, 1, 6].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F4F6F8"));
                Ausentismos.Cells[1, 1, 1, 6].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                Ausentismos.Cells[1, 1, 1, 6].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                Ausentismos.Cells[1, 1, 1, 6].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                Ausentismos.Cells[1, 1, 1, 6].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                Ausentismos.Cells[1, 1, 1, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Ausentismos.Cells[1, 1, 1, 6].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                Ausentismos.Cells[1, 1, 1, 6].Style.Font.Bold = true;
                string FechaIniMensaje = ResultsDatesPaysheet.DateStart;
                var validationDate1 = Ausentismos.Cells["C2:C500"].DataValidation.AddDateTimeDataValidation();
                validationDate1.ErrorStyle = ExcelDataValidationWarningStyle.stop;
                validationDate1.PromptTitle = "Registro de Ausentismo";
                validationDate1.ErrorTitle = "La fecha ingresada no es válida";
                validationDate1.Operator = ExcelDataValidationOperator.between;
                validationDate1.Formula.Value = Convert.ToDateTime(FechaIniMensaje);
                validationDate1.Formula2.Value = Convert.ToDateTime(ResultsDatesPaysheet.DateEnd);
                validationDate1.Prompt = "El valor máximo disponible es de " + FechaIniMensaje + " a " + ResultsDatesPaysheet.DateEnd + " Nota: Si no desea ingresar algún valor, use la tecla SUPR";
                validationDate1.Error = "Solo puede ingresar valores entre " + FechaIniMensaje + " y " + ResultsDatesPaysheet.DateEnd;
                validationDate1.ShowInputMessage = true;
                validationDate1.ShowErrorMessage = true;
                var validationDate2 = Ausentismos.Cells["D2:D500"].DataValidation.AddDateTimeDataValidation();
                validationDate2.ErrorStyle = ExcelDataValidationWarningStyle.stop;
                validationDate2.PromptTitle = "Registro de Ausentismo";
                validationDate2.ErrorTitle = "La fecha ingresada no es válida";
                validationDate2.Operator = ExcelDataValidationOperator.between;
                validationDate2.Formula.Value = Convert.ToDateTime(FechaIniMensaje);
                validationDate2.Formula2.Value = Convert.ToDateTime("31/12/" + DateTime.Now.Year);
                validationDate2.Prompt = "El valor máximo disponible es de " + FechaIniMensaje + " a " + "31/12/" + DateTime.Now.Year + " Nota: Si no desea ingresar algún valor, use la tecla SUPR";
                validationDate2.Error = "Solo puede ingresar valores entre " + FechaIniMensaje + " y " + "31/12/" + DateTime.Now.Year;
                validationDate2.ShowInputMessage = true;
                validationDate2.ShowErrorMessage = true;
                var validationDate4 = Ausentismos.Cells["E2:E500"].DataValidation.AddDateTimeDataValidation();
                validationDate4.ErrorStyle = ExcelDataValidationWarningStyle.stop;
                validationDate4.PromptTitle = "Registro de Ausentismo";
                validationDate4.ErrorTitle = "La fecha ingresada no es válida";
                validationDate4.Operator = ExcelDataValidationOperator.between;
                validationDate4.Formula.Value = Convert.ToDateTime("01/01/" + DateTime.Now.Year);
                validationDate4.Formula2.Value = Convert.ToDateTime("31/12/" + DateTime.Now.Year);
                validationDate4.Prompt = "El valor máximo disponible es de " + "01/01/" + DateTime.Now.Year + " a " + "31/12/" + DateTime.Now.Year + " Nota: Si no desea ingresar algún valor, use la tecla SUPR";
                validationDate4.Error = "Solo puede ingresar valores entre " + "01/01/" + DateTime.Now.Year + " y " + "31/12/" + DateTime.Now.Year;
                validationDate4.ShowInputMessage = true;
                validationDate4.ShowErrorMessage = true;
                var validationDate5 = Ausentismos.Cells["F2:F500"].DataValidation.AddDateTimeDataValidation();
                validationDate5.ErrorStyle = ExcelDataValidationWarningStyle.stop;
                validationDate5.PromptTitle = "Registro de Ausentismo";
                validationDate5.ErrorTitle = "La fecha ingresada no es válida";
                validationDate5.Operator = ExcelDataValidationOperator.between;
                validationDate5.Formula.Value = Convert.ToDateTime("01/01/" + DateTime.Now.Year);
                validationDate5.Formula2.Value = Convert.ToDateTime("31/12/" + DateTime.Now.Year);
                validationDate5.Prompt = "El valor máximo disponible es de " + "01/01/" + DateTime.Now.Year + " a " + "31/12/" + DateTime.Now.Year + " Nota: Si no desea ingresar algún valor, use la tecla SUPR";
                validationDate5.Error = "Solo puede ingresar valores entre " + "01/01/" + DateTime.Now.Year + " y " + "31/12/" + DateTime.Now.Year;
                validationDate5.ShowInputMessage = true;
                validationDate5.ShowErrorMessage = true;
                Ausentismos.Cells[Ausentismos.Dimension.Address].Style.Font.Name = "Century Gothic";
                Ausentismos.Cells["A2:A500"].Style.Locked = false;
                Ausentismos.Cells["B2:B500"].Style.Locked = false;
                Ausentismos.Cells["C2:C500"].Style.Locked = false;
                Ausentismos.Cells["D2:D500"].Style.Locked = false;
                Ausentismos.Cells["E2:E500"].Style.Locked = false;
                Ausentismos.Cells["F2:F500"].Style.Locked = false;
                Ausentismos.Cells[Ausentismos.Dimension.Address].AutoFitColumns();
                Ausentismos.Column(1).Width = 50;
                Ausentismos.Column(2).Width = 30;
                //Hoja de retirtos
                Retiros.Protection.IsProtected = true;
                Retiros.Cells["A1"].Value = "DOCUMENTO";
                Retiros.Cells["A1"].Style.Font.Size = 13;
                Retiros.Cells["B1"].Value = "EMPLEADO";
                Retiros.Cells["B1"].Style.Font.Size = 13;
                Retiros.Cells["C1"].Value = "FECHA DE RETIRO";
                Retiros.Cells["C1"].Style.Font.Size = 13;
                Retiros.Cells["D1"].Value = "CAUSA DE RETIRO";
                Retiros.Cells["D1"].Style.Font.Size = 13;
                filas = 2;
                foreach (var item in ListCauseRetreat)
                {
                    Formulas.Cells[filas, 3].Value = item.cau_ret.Trim() + " - " + item.nom_ret;
                    filas++;
                }
                var dd3 = Retiros.Cells["D2:D500"].DataValidation.AddListDataValidation() as ExcelDataValidationList;
                dd3.AllowBlank = true;
                dd3.Formula.ExcelFormula = "Formulas!C$2:C$" + (ListCauseRetreat.Count + 1);
                Retiros.Row(1).Height = 41.25;
                Retiros.Cells[1, 1, 1, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                Retiros.Cells[1, 1, 1, 4].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F4F6F8"));
                Retiros.Cells[1, 1, 1, 4].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                Retiros.Cells[1, 1, 1, 4].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                Retiros.Cells[1, 1, 1, 4].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                Retiros.Cells[1, 1, 1, 4].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                Retiros.Cells[1, 1, 1, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Retiros.Cells[1, 1, 1, 4].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                Retiros.Cells[1, 1, 1, 4].Style.Font.Bold = true;


                Retiros.Cells[Retiros.Dimension.Address].Style.Font.Name = "Century Gothic";
                Retiros.Cells[Retiros.Dimension.Address].AutoFitColumns();
                Formulas.Hidden = eWorkSheetHidden.Hidden;
                Retiros.Cells["C2:C" + (FilasRetiros - 1)].Style.Locked = false;
                Retiros.Cells["D2:D500"].Style.Locked = false;
                hc.GuardarLogs("El cliente ha generado el archivo de novedades de nómina para el corte actual",
                               "Proceso correcto, generación de archivo correcta",
                               User.Identity.Name,
                               Session["NIT"].ToString());
                return File(Excel.GetAsByteArray(), "application/octet-stream", "Plantilla novedades " + sd.Date + ".xlsx");
            }
            catch (Exception ex)
            {
                hc.GuardarLogs("Error al generar en la generación del excel.", ex.Message, NameCompany(), Session["NIT"].ToString());
                Alerts("No pudimos generar el archivo.", NotificationTypes.error);
                return RedirectToAction("NoveltiesAndAbsences");
            }
        }
        [HttpPost]
        public ActionResult ValidarArchivo(RuteToCreate RuteCreate, HttpPostedFileBase File)
        {
            try
            {
                //Se generan las variables de session
                if (!SessionApi())
                {
                    hc.GuardarLogs("Variables de session NO generadas desde el metodo ValidarArchivo en el controlador Transactions.",
                                   "Proceso incompleto, no se pudieron generar las variables de sesión para el proceso",
                                   User.Identity.Name,
                                   "N/A");
                    return View("ErrorApp");
                }
                //Validaciones generales
                if (File == null)
                {
                    hc.GuardarLogs("El cliente ha intentado guardar el archivo de novedades sin entregar un archivo",
                                          "Proceso incompleto, se debe entregar un archivo",
                                          NameCompany(),
                                          Session["NIT"].ToString());
                    Alerts("Debes agregar un archivo.", NotificationTypes.warning);
                    return RedirectToAction("NoveltiesAndAbsences",
                                             new ToPrincipalView
                                             {
                                                 date = RuteCreate.Date,
                                                 Agree = RuteCreate.Agre,
                                                 IncomeType = RuteCreate.IncomeType
                                             });
                }
                string extension = Path.GetExtension(File.FileName);
                if (extension != ".xlsx")
                {
                    hc.GuardarLogs("El cliente ha intentado guardar el archivo de novedades con una extensión o tio de archivo incompatible",
                                          "Proceso incompleto, para ingresar un archivo debe ser un archivo de tipo xlsx",
                                          NameCompany(),
                                          Session["NIT"].ToString());
                    Alerts("Este archivo no es compatible con la extensión original.", NotificationTypes.warning);
                    return RedirectToAction("NoveltiesAndAbsences",
                                            new ToPrincipalView
                                            {
                                                date = RuteCreate.Date,
                                                Agree = RuteCreate.Agre,
                                                IncomeType = RuteCreate.IncomeType
                                            });
                }
                //Hago creacipon del archivo de respuesta
                IRestResponse response;
                NoveltyData noveltyData = new NoveltyData();
                noveltyData.Date = RuteCreate.Date.Trim();
                noveltyData.Agreement = RuteCreate.Agre.Trim();
                response = new ApiUrl().UrlExecute("PaysheetEntryAndSaveErrorsExcel/DeletePaysheetEntryErrors", Session["ApiToken"].ToString(), "POST", JsonConvert.SerializeObject(noveltyData));
                if (response.StatusDescription == "OK")
                {
                    ResponseJson RJson = JsonConvert.DeserializeObject<ResponseJson>(response.Content);
                    if (!RJson.succeeded)
                    {
                        Alerts("Tenemos unos inconvenientes con lel proceso de revisón de ingreso. Inténtalo más tarde", NotificationTypes.warning);
                        hc.GuardarLogs("No podemos borrar los datos de las entradas y datos de guardados del excel de novedades", RJson.Errors[0].Description, NameCompany(), Session["NIT"].ToString());
                        return RedirectToAction("NoveltiesAndAbsences", new ToPrincipalView { date = RuteCreate.Date, Agree = RuteCreate.Agre, IncomeType = RuteCreate.IncomeType });
                    }
                }
                else
                {
                    Alerts("Estamos presentando inconvenientes en la comunicación en este momento, intenta nuevamente en unos minutos.", NotificationTypes.error);
                    hc.GuardarLogs("No podemos borrar los datos de las entradas y datos de guardados del excel de novedades. APi responde de forma inesperada", response.StatusDescription, NameCompany(), Session["NIT"].ToString());
                    return RedirectToAction("NoveltiesAndAbsences", new ToPrincipalView
                    {
                        date = RuteCreate.Date,
                        Agree = RuteCreate.Agre,
                        IncomeType = RuteCreate.IncomeType
                    });
                }
                int RegisterNovelties = 0;
                int RegisterAusentsims = 0;
                int RegisterRetreats = 0;
                bool ContinueValidate = true;
                List<PaysheetEntryAndSaveErrorsExcel> SaveErrorsExcel = new List<PaysheetEntryAndSaveErrorsExcel>();
                PaysheetEntryAndSaveErrorsExcel ErrorEntry = new PaysheetEntryAndSaveErrorsExcel();
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelPackage Archivo = new ExcelPackage(File.InputStream);
                ExcelWorksheet Novedades = Archivo.Workbook.Worksheets[0];
                ExcelWorksheet Ausentismos = Archivo.Workbook.Worksheets[1];
                ExcelWorksheet Retiros = Archivo.Workbook.Worksheets[2];
                //...NOVEDADES MAL INGRESADAS...
                List<NoveltiesData> ListNovedades = new List<NoveltiesData>();
                NoveltiesData Novedad = new NoveltiesData();
                var Conceptsrow = Novedades.Cells[1, 14, 1, Novedades.Dimension.End.Column];
                string[] Conceptos = new string[2];
                int LentValue = 0;
                for (int i = 2; i <= Novedades.Dimension.End.Row; i++)
                {
                    var row = Novedades.Cells[i, 1, i, Novedades.Dimension.End.Column];
                    if (row["A" + i].Value == null || row["B" + i].Value == null)
                    {
                        break;
                    }
                    for (int j = 14; j <= Novedades.Dimension.End.Column; j++)
                    {
                        try
                        {
                            if (row[i, j].Value != null && !string.IsNullOrEmpty(row[i, j].Value.ToString().Trim()) && row[i, j].Value.ToString().Trim() != "0")
                            {
                                Novedad.Id = 0;
                                Novedad.Nit = Session["NIT"].ToString();
                                Novedad.DateCut = RuteCreate.Date.Trim();
                                Novedad.Employee_Id = row["A" + i].Value.ToString();
                                LentValue = row[i, j].Value.ToString().Length;
                                Novedad.Value = row[i, j].Value.ToString().Trim();
                                Conceptos = Conceptsrow[1, j].Value.ToString().Split('-');
                                Novedad.Concept_Id = Conceptos[0].Trim();
                                Novedad.Agreement = RuteCreate.Agre.Trim();
                                Novedad.Flag = 1;
                                Novedad.Concept_Name = Conceptos[1].Substring(1);
                                ListNovedades.Add(Novedad);
                                Novedad = new NoveltiesData();
                                Conceptos = new string[2];
                                RegisterNovelties = RegisterNovelties + 1;
                            }
                        }
                        catch
                        {
                            ErrorEntry.Id = 0;
                            ErrorEntry.Nit = Session["NIT"].ToString();
                            ErrorEntry.Date_Cut = RuteCreate.Date.Trim();
                            ErrorEntry.Agreement = RuteCreate.Agre.Trim();
                            ErrorEntry.Flag = 1;
                            if (Conceptsrow[1, j].Value != null)
                            {
                                Conceptos = Conceptsrow[1, j].Value.ToString().Split('-');
                                ErrorEntry.Concept_Id = Conceptos[0].Trim();
                                ErrorEntry.Concept_Name = Conceptos[1].Substring(1);

                            }
                            else
                            {
                                ErrorEntry.Concept_Id = "No ingresado";
                                ErrorEntry.Concept_Name = "No ingresado";
                            }
                            if (row["A" + i].Value != null)
                            {
                                ErrorEntry.Employee = row["A" + i].Value.ToString();
                            }
                            else
                            {
                                ErrorEntry.Employee = "No ingresado";
                            }
                            if (row[i, j].Value == null)
                            {
                                ErrorEntry.Values = row[i, j].Value.ToString().Trim();
                            }
                            else
                            {
                                ErrorEntry.Values = "No ingresado";
                            }
                            ErrorEntry.MessageError = "La novedad posee datos errados o incompletos que no permiten analizarla, recuerde guardar los conceptos y formas del archico cuando se le entrego.";
                            SaveErrorsExcel.Add(ErrorEntry);
                            ErrorEntry = new PaysheetEntryAndSaveErrorsExcel();
                            RegisterNovelties = RegisterNovelties + 1;
                        }
                    }
                }
                //Valido que en la lista los valores todos sean de tipo decimal y que no contengan valores extraños.
                foreach (var Novelt in ListNovedades)
                {
                    ContinueValidate = true;
                    if (Novelt.Value.Contains(".") && ContinueValidate)
                    {
                        ErrorEntry.Id = 0;
                        ErrorEntry.Nit = Novelt.Nit;
                        ErrorEntry.Date_Cut = Novelt.DateCut.Trim();
                        ErrorEntry.Agreement = Novelt.Agreement.Trim();
                        ErrorEntry.Flag = 1;
                        ErrorEntry.Employee = Novelt.Employee_Id;
                        ErrorEntry.Concept_Id = Novelt.Concept_Id;
                        ErrorEntry.Concept_Name = Novelt.Concept_Name;
                        ErrorEntry.Values = Novelt.Value;
                        ErrorEntry.MessageError = "El valor que esta intentando guardar no es valido, los separadores son comas (,)";
                        SaveErrorsExcel.Add(ErrorEntry);
                        ErrorEntry = new PaysheetEntryAndSaveErrorsExcel();
                        ContinueValidate = false;
                    }
                    if (string.IsNullOrEmpty(Novelt.Value) && ContinueValidate)
                    {
                        ErrorEntry.Id = 0;
                        ErrorEntry.Nit = Novelt.Nit;
                        ErrorEntry.Date_Cut = Novelt.DateCut.Trim();
                        ErrorEntry.Agreement = Novelt.Agreement.Trim();
                        ErrorEntry.Flag = 1;
                        ErrorEntry.Employee = Novelt.Employee_Id;
                        ErrorEntry.Concept_Id = Novelt.Concept_Id;
                        ErrorEntry.Concept_Name = Novelt.Concept_Name;
                        ErrorEntry.Values = Novelt.Value;
                        ErrorEntry.MessageError = "El valor que esta intentando guardar no es valido, debe ingresar algun númerico";
                        SaveErrorsExcel.Add(ErrorEntry);
                        ErrorEntry = new PaysheetEntryAndSaveErrorsExcel();
                        ContinueValidate = false;
                    }
                    if (ContinueValidate)
                    {
                        try
                        {
                            Convert.ToDecimal(Novelt.Value);
                        }
                        catch
                        {
                            ErrorEntry.Id = 0;
                            ErrorEntry.Nit = Novelt.Nit;
                            ErrorEntry.Date_Cut = Novelt.DateCut.Trim();
                            ErrorEntry.Agreement = Novelt.Agreement.Trim();
                            ErrorEntry.Flag = 1;
                            ErrorEntry.Employee = Novelt.Employee_Id;
                            ErrorEntry.Concept_Id = Novelt.Concept_Id;
                            ErrorEntry.Concept_Name = Novelt.Concept_Name;
                            ErrorEntry.Values = Novelt.Value;
                            ErrorEntry.MessageError = "El valor ingresado para la novedad no es valido. solo se pueden ingresar numeros con dos decimales separados por comas (0,00)";
                            SaveErrorsExcel.Add(ErrorEntry);
                            ErrorEntry = new PaysheetEntryAndSaveErrorsExcel();
                            ContinueValidate = false;
                        }
                    }
                }
                //...AUSENTISMOS MAL INGRESADOS...
                List<AbsencesData> ListAusetismos = new List<AbsencesData>();
                AbsencesData Ausentismo = new AbsencesData();
                string[] Empleado = new string[2];
                string[] TipoAusentismo = new string[2];
                for (int i = 2; i <= Ausentismos.Dimension.End.Row; i++)
                {
                    var row = Ausentismos.Cells[i, 1, i, Ausentismos.Dimension.End.Column];
                    if (row["A" + i].Value == null)
                    {
                        break;
                    }
                    try
                    {
                        Ausentismo.Id = 0;
                        Ausentismo.Nit = Session["NIT"].ToString();
                        if (row["D" + i].Value == null)
                        {
                            Ausentismo.DateEnd = "No ingresado";
                        }
                        else
                        {
                            Ausentismo.DateEnd = row["D" + i].Value.ToString().Substring(0, 10);
                            if (Ausentismo.DateEnd.Trim().Length == 9)
                            {
                                Ausentismo.DateEnd = "0" + Ausentismo.DateEnd.Trim();
                            }
                        }
                        Ausentismo.DateCut = RuteCreate.Date;
                        Empleado = (row["A" + i].Value.ToString()).Split('-');
                        Ausentismo.Employee_Id = Empleado[0].Trim();
                        if (row["C" + i].Value == null)
                        {
                            Ausentismo.DateStart = "No ingresado";
                        }
                        else
                        {
                            Ausentismo.DateStart = row["C" + i].Value.ToString().Substring(0, 10);
                            if (Ausentismo.DateStart.Trim().Length == 9)
                            {
                                Ausentismo.DateStart = "0" + Ausentismo.DateStart.Trim();
                            }
                        }
                        if (row["E" + i].Value == null)
                        {
                            Ausentismo.DateReallyStart = "No ingresado";
                        }
                        else
                        {
                            Ausentismo.DateReallyStart = row["E" + i].Value.ToString().Substring(0, 10);
                            if (Ausentismo.DateReallyStart.Trim().Length == 9)
                            {
                                Ausentismo.DateReallyStart = "0" + Ausentismo.DateReallyStart.Trim();
                            }
                        }
                        if (row["F" + i].Value == null)
                        {
                            Ausentismo.DateReallyEnd = "No ingresado";
                        }
                        else
                        {
                            Ausentismo.DateReallyEnd = row["F" + i].Value.ToString().Substring(0, 10);
                            if (Ausentismo.DateReallyEnd.Trim().Length == 9)
                            {
                                Ausentismo.DateReallyEnd = "0" + Ausentismo.DateReallyEnd.Trim();
                            }
                        }
                        Ausentismo.Agreement = RuteCreate.Agre;
                        Ausentismo.Flag = 2;
                        TipoAusentismo = (row["B" + i].Value.ToString()).Split('-');
                        Ausentismo.Absenteeism_Id = TipoAusentismo[0].Trim();
                        Ausentismo.Absenteeism_Name = TipoAusentismo[1].Substring(1);
                        ListAusetismos.Add(Ausentismo);
                        Empleado = new string[2];
                        TipoAusentismo = new string[2];
                        Ausentismo = new AbsencesData();
                        RegisterAusentsims = RegisterAusentsims + 1;
                    }
                    catch
                    {
                        ErrorEntry.Id = 0;
                        ErrorEntry.Nit = Session["NIT"].ToString();
                        ErrorEntry.Date_Cut = RuteCreate.Date.Trim();
                        ErrorEntry.Agreement = RuteCreate.Agre.Trim();
                        ErrorEntry.Flag = 2;
                        if (row["A" + i].Value != null)
                        {
                            ErrorEntry.Employee = row["A" + i].Value.ToString();
                        }
                        else
                        {
                            ErrorEntry.Employee = "No ingresado";
                        }
                        if (row["B" + i].Value != null)
                        {
                            ErrorEntry.Absenteeism_Id = row["B" + i].Value.ToString();
                            ErrorEntry.Absenteeism_Name = row["B" + i].Value.ToString();
                        }
                        else
                        {
                            ErrorEntry.Absenteeism_Id = "No ingresado";
                            ErrorEntry.Absenteeism_Name = "No ingresado";
                        }
                        if (row["C" + i].Value != null)
                        {
                            ErrorEntry.DateStart = row["C" + i].Value.ToString();
                        }
                        else
                        {
                            ErrorEntry.DateStart = "No ingresado";
                        }
                        if (row["D" + i].Value != null)
                        {
                            ErrorEntry.DateEnd = row["D" + i].Value.ToString();
                        }
                        else
                        {
                            ErrorEntry.DateEnd = "No ingresado";
                        }
                        if (row["E" + i].Value != null)
                        {
                            ErrorEntry.DateReallyStart = row["E" + i].Value.ToString();
                        }
                        else
                        {
                            ErrorEntry.DateReallyStart = "No ingresado";
                        }
                        if (row["F" + i].Value != null)
                        {
                            ErrorEntry.DateReallyEnd = row["F" + i].Value.ToString();
                        }
                        else
                        {
                            ErrorEntry.DateReallyEnd = "No ingresado";
                        }
                        ErrorEntry.MessageError = "La información de este ausentismo no fue ingresada correctamente, le faltan datos no cumple con los estandares de ingreso";
                        SaveErrorsExcel.Add(ErrorEntry);
                        ErrorEntry = new PaysheetEntryAndSaveErrorsExcel();
                        RegisterAusentsims = RegisterAusentsims + 1;
                    }
                }
                foreach (var Ausen in ListAusetismos)
                {
                    ContinueValidate = true;
                    if (ContinueValidate)
                    {
                        if (string.IsNullOrEmpty(Ausen.DateStart) ||
                            string.IsNullOrEmpty(Ausen.DateEnd) ||
                            string.IsNullOrEmpty(Ausen.DateReallyStart) ||
                            string.IsNullOrEmpty(Ausen.DateReallyEnd))
                        {
                            ErrorEntry.Id = 0;
                            ErrorEntry.Nit = Ausen.Nit.Trim();
                            ErrorEntry.Date_Cut = Ausen.DateCut.Trim();
                            ErrorEntry.Agreement = Ausen.Agreement.Trim();
                            ErrorEntry.Flag = 2;
                            ErrorEntry.Employee = Ausen.Employee_Id;
                            ErrorEntry.Absenteeism_Id = Ausen.Absenteeism_Id;
                            ErrorEntry.Absenteeism_Name = Ausen.Absenteeism_Name;
                            ErrorEntry.DateStart = Ausen.DateStart;
                            ErrorEntry.DateEnd = Ausen.DateEnd;
                            ErrorEntry.DateReallyStart = Ausen.DateReallyStart;
                            ErrorEntry.DateReallyEnd = Ausen.DateReallyEnd;
                            ErrorEntry.MessageError = "Los datos que intenta ingresar estan incompletos, debe ingresar todas las fechas de registro y fechas reales ";
                            SaveErrorsExcel.Add(ErrorEntry);
                            ErrorEntry = new PaysheetEntryAndSaveErrorsExcel();
                            ContinueValidate = false;
                        }
                    }


                    if (string.IsNullOrEmpty(Ausen.Absenteeism_Id) && ContinueValidate)
                    {
                        ErrorEntry.Id = 0;
                        ErrorEntry.Nit = Ausen.Nit.Trim();
                        ErrorEntry.Date_Cut = Ausen.DateCut.Trim();
                        ErrorEntry.Agreement = Ausen.Agreement.Trim();
                        ErrorEntry.Flag = 2;
                        ErrorEntry.Employee = Ausen.Employee_Id;
                        ErrorEntry.Absenteeism_Id = Ausen.Absenteeism_Id;
                        ErrorEntry.Absenteeism_Name = Ausen.Absenteeism_Name;
                        ErrorEntry.DateStart = Ausen.DateStart;
                        ErrorEntry.DateEnd = Ausen.DateEnd;
                        ErrorEntry.DateReallyStart = Ausen.DateReallyStart;
                        ErrorEntry.DateReallyEnd = Ausen.DateReallyEnd;
                        ErrorEntry.MessageError = "Se debe ingresar los datos del tipo de ausentismo respetando el estandar codigo - Nombre";
                        SaveErrorsExcel.Add(ErrorEntry);
                        ErrorEntry = new PaysheetEntryAndSaveErrorsExcel();
                        ContinueValidate = false;
                    }
                    if (ContinueValidate)
                    {
                        try
                        {
                            Convert.ToDateTime(Ausen.DateStart);
                        }
                        catch
                        {
                            ErrorEntry.Id = 0;
                            ErrorEntry.Nit = Ausen.Nit.Trim();
                            ErrorEntry.Date_Cut = Ausen.DateCut.Trim();
                            ErrorEntry.Agreement = Ausen.Agreement.Trim();
                            ErrorEntry.Flag = 2;
                            ErrorEntry.Employee = Ausen.Employee_Id;
                            ErrorEntry.Absenteeism_Id = Ausen.Absenteeism_Id;
                            ErrorEntry.Absenteeism_Name = Ausen.Absenteeism_Name;
                            ErrorEntry.DateStart = Ausen.DateStart;
                            ErrorEntry.DateEnd = Ausen.DateEnd;
                            ErrorEntry.DateReallyStart = Ausen.DateReallyStart;
                            ErrorEntry.DateReallyEnd = Ausen.DateReallyEnd;
                            ErrorEntry.MessageError = "El dato que intenta ingresar como fecha de registro inicial no es valido";
                            SaveErrorsExcel.Add(ErrorEntry);
                            ErrorEntry = new PaysheetEntryAndSaveErrorsExcel();
                            ContinueValidate = false;
                        }
                    }

                    if (ContinueValidate)
                    {
                        try
                        {
                            Convert.ToDateTime(Ausen.DateEnd);
                        }
                        catch
                        {
                            ErrorEntry.Id = 0;
                            ErrorEntry.Nit = Ausen.Nit.Trim();
                            ErrorEntry.Date_Cut = Ausen.DateCut.Trim();
                            ErrorEntry.Agreement = Ausen.Agreement.Trim();
                            ErrorEntry.Flag = 2;
                            ErrorEntry.Employee = Ausen.Employee_Id;
                            ErrorEntry.Absenteeism_Id = Ausen.Absenteeism_Id;
                            ErrorEntry.Absenteeism_Name = Ausen.Absenteeism_Name;
                            ErrorEntry.DateStart = Ausen.DateStart;
                            ErrorEntry.DateEnd = Ausen.DateEnd;
                            ErrorEntry.DateReallyStart = Ausen.DateReallyStart;
                            ErrorEntry.DateReallyEnd = Ausen.DateReallyEnd;
                            ErrorEntry.MessageError = "El dato que intenta ingresar como fecha final de registro no es valido";
                            SaveErrorsExcel.Add(ErrorEntry);
                            ErrorEntry = new PaysheetEntryAndSaveErrorsExcel();
                            ContinueValidate = false;
                        }
                    }
                    if (ContinueValidate)
                    {
                        try
                        {
                            Convert.ToDateTime(Ausen.DateReallyStart);
                        }
                        catch
                        {
                            ErrorEntry.Id = 0;
                            ErrorEntry.Nit = Ausen.Nit.Trim();
                            ErrorEntry.Date_Cut = Ausen.DateCut.Trim();
                            ErrorEntry.Agreement = Ausen.Agreement.Trim();
                            ErrorEntry.Flag = 2;
                            ErrorEntry.Employee = Ausen.Employee_Id;
                            ErrorEntry.Absenteeism_Id = Ausen.Absenteeism_Id;
                            ErrorEntry.Absenteeism_Name = Ausen.Absenteeism_Name;
                            ErrorEntry.DateStart = Ausen.DateStart;
                            ErrorEntry.DateEnd = Ausen.DateEnd;
                            ErrorEntry.DateReallyStart = Ausen.DateReallyStart;
                            ErrorEntry.DateReallyEnd = Ausen.DateReallyEnd;
                            ErrorEntry.MessageError = "El dato que intenta guardar como fecha real de inicio de ausentimo no es valido";
                            SaveErrorsExcel.Add(ErrorEntry);
                            ErrorEntry = new PaysheetEntryAndSaveErrorsExcel();
                            ContinueValidate = false;
                        }
                    }
                    if (ContinueValidate)
                    {
                        try
                        {
                            Convert.ToDateTime(Ausen.DateReallyEnd);
                        }
                        catch
                        {
                            ErrorEntry.Id = 0;
                            ErrorEntry.Nit = Ausen.Nit.Trim();
                            ErrorEntry.Date_Cut = Ausen.DateCut.Trim();
                            ErrorEntry.Agreement = Ausen.Agreement.Trim();
                            ErrorEntry.Flag = 2;
                            ErrorEntry.Employee = Ausen.Employee_Id;
                            ErrorEntry.Absenteeism_Id = Ausen.Absenteeism_Id;
                            ErrorEntry.Absenteeism_Name = Ausen.Absenteeism_Name;
                            ErrorEntry.DateStart = Ausen.DateStart;
                            ErrorEntry.DateEnd = Ausen.DateEnd;
                            ErrorEntry.DateReallyStart = Ausen.DateReallyStart;
                            ErrorEntry.DateReallyEnd = Ausen.DateReallyEnd;
                            ErrorEntry.MessageError = "El dato que intenta guardar como real de final de ausentismo no es valido";
                            SaveErrorsExcel.Add(ErrorEntry);
                            ErrorEntry = new PaysheetEntryAndSaveErrorsExcel();
                            ContinueValidate = false;
                        }
                    }

                }
                //...RETIROS MAL INGRESADAS...
                List<RetreatsData> ListRetiros = new List<RetreatsData>();
                RetreatsData Retiro = new RetreatsData();
                string[] Causa = new string[2];
                for (int i = 2; i <= Retiros.Dimension.End.Row; i++)
                {
                    var row = Retiros.Cells[i, 1, i, Retiros.Dimension.End.Column];
                    if (row["A" + i].Value == null || row["B" + i].Value == null)
                    {
                        break;
                    }
                    try
                    {
                        if (row["C" + i].Value != null)
                        {
                            Retiro.Id = 0;
                            Retiro.Nit = Session["NIT"].ToString();
                            Retiro.DateEnd = row["C" + i].Value.ToString().Substring(0, 10);
                            if (Retiro.DateEnd.Trim().Length == 9)
                            {
                                Retiro.DateEnd = "0" + Retiro.DateEnd.Trim();
                            }
                            Retiro.DateCut = RuteCreate.Date;
                            Retiro.Employee_Id = row["A" + i].Value.ToString().Trim();
                            Retiro.Agreement = RuteCreate.Agre;
                            Retiro.Flag = 3;
                            Causa = (row["D" + i].Value.ToString()).Split('-');
                            Retiro.cau_ret = Causa[0].Trim();
                            ListRetiros.Add(Retiro);
                            Retiro = new RetreatsData();
                            RegisterRetreats = RegisterRetreats + 1;
                        }
                    }
                    catch
                    {
                        ErrorEntry.Id = 0;
                        ErrorEntry.Nit = Session["NIT"].ToString();
                        ErrorEntry.Date_Cut = RuteCreate.Date.Trim();
                        ErrorEntry.Agreement = RuteCreate.Agre.Trim();
                        ErrorEntry.Flag = 3;
                        if (row["A" + i].Value != null)
                        {
                            ErrorEntry.Employee = row["A" + i].Value.ToString();
                        }
                        else
                        {
                            ErrorEntry.Employee = "No ingresado";
                        }
                        if (row["C" + i].Value != null)
                        {
                            ErrorEntry.DateEnd = row["C" + i].Value.ToString();
                        }
                        else
                        {
                            ErrorEntry.DateEnd = "No ingresado";
                        }
                        ErrorEntry.MessageError = "La información de este retiro no fue ingresada correctamente, le faltan datos no cumple con los estandares de ingreso";
                        SaveErrorsExcel.Add(ErrorEntry);
                        ErrorEntry = new PaysheetEntryAndSaveErrorsExcel();
                    }
                }
                foreach (var Ret in ListRetiros)
                {
                    try
                    {
                        Convert.ToDateTime(Ret.DateEnd);
                    }
                    catch
                    {
                        ErrorEntry.Id = 0;
                        ErrorEntry.Nit = Ret.Nit;
                        ErrorEntry.Date_Cut = Ret.DateCut;
                        ErrorEntry.Agreement = Ret.Agreement;
                        ErrorEntry.Flag = 3;
                        ErrorEntry.Employee = Ret.Employee_Id;
                        ErrorEntry.DateEnd = Ret.DateEnd;
                        ErrorEntry.MessageError = "El dato que intenta ingresar como fecha de retiro no es valido";
                        SaveErrorsExcel.Add(ErrorEntry);
                        ErrorEntry = new PaysheetEntryAndSaveErrorsExcel();
                        ContinueValidate = false;
                    }
                }
                ErrorEntry.Id = 0;
                ErrorEntry.Nit = Session["NIT"].ToString();
                ErrorEntry.Date_Cut = RuteCreate.Date.Trim();
                ErrorEntry.Agreement = RuteCreate.Agre.Trim();
                ErrorEntry.Flag = 1;
                ErrorEntry.Employee = "N/A";
                ErrorEntry.Concept_Id = "N/A";
                ErrorEntry.Concept_Name = "Registro total de novedades de devengos y deducciones accesibles por el archivo";
                ErrorEntry.Values = RegisterNovelties.ToString();
                SaveErrorsExcel.Add(ErrorEntry);
                ErrorEntry = new PaysheetEntryAndSaveErrorsExcel();
                ErrorEntry.Id = 0;
                ErrorEntry.Nit = Session["NIT"].ToString();
                ErrorEntry.Date_Cut = RuteCreate.Date.Trim();
                ErrorEntry.Agreement = RuteCreate.Agre.Trim();
                ErrorEntry.Flag = 1;
                ErrorEntry.Employee = "N/A";
                ErrorEntry.Concept_Id = "N/A";
                ErrorEntry.Concept_Name = "Registro total de ausentismos accesibles por el archivo";
                ErrorEntry.Values = RegisterAusentsims.ToString();
                SaveErrorsExcel.Add(ErrorEntry);
                ErrorEntry = new PaysheetEntryAndSaveErrorsExcel();
                ErrorEntry.Id = 0;
                ErrorEntry.Nit = Session["NIT"].ToString();
                ErrorEntry.Date_Cut = RuteCreate.Date.Trim();
                ErrorEntry.Agreement = RuteCreate.Agre.Trim();
                ErrorEntry.Flag = 1;
                ErrorEntry.Employee = "N/A";
                ErrorEntry.Concept_Id = "N/A";
                ErrorEntry.Concept_Name = "Registro total de retiros accesibles por el archivo";
                ErrorEntry.Values = RegisterRetreats.ToString();
                SaveErrorsExcel.Add(ErrorEntry);
                ErrorEntry = new PaysheetEntryAndSaveErrorsExcel();
                int ResultsInserts = 0;
                //Guardamos listas de incosistencias
                response = new ApiUrl().UrlExecute("PaysheetEntryAndSaveErrorsExcel/SaveListPaysheetEntryErrors",
                                                    Session["ApiToken"].ToString(),
                                                    "POST",
                                                    JsonConvert.SerializeObject(SaveErrorsExcel));
                if (response.StatusDescription == "OK")
                {
                    ResultsInserts = 2;
                    ResponseJson RJson = JsonConvert.DeserializeObject<ResponseJson>(response.Content);
                    if (!RJson.succeeded)
                    {
                        Alerts("Tenemos unos inconvenientes con el proceso de revisón de ingreso. Inténtalo más tarde", NotificationTypes.warning);
                        hc.GuardarLogs("No pudimos guardar las novedades del archivo",
                                        RJson.Errors[0].Description,
                                        NameCompany(),
                                        Session["NIT"].ToString());
                        return RedirectToAction("NoveltiesAndAbsences",
                                                new ToPrincipalView
                                                {
                                                    date = RuteCreate.Date,
                                                    Agree = RuteCreate.Agre,
                                                    IncomeType = RuteCreate.IncomeType
                                                });
                    }
                }
                else
                {
                    Alerts("Estamos presentando inconvenientes en la comunicación en este momento, intenta nuevamente en unos minutos.", NotificationTypes.error);
                    hc.GuardarLogs("No podemos guardar los datos de las entradas y datos de guardados del excel de novedades. APi responde de forma inesperada",
                                    response.StatusDescription,
                                    NameCompany(),
                                    Session["NIT"].ToString());
                    return RedirectToAction("NoveltiesAndAbsences", new ToPrincipalView
                    {
                        date = RuteCreate.Date,
                        Agree = RuteCreate.Agre,
                        IncomeType = RuteCreate.IncomeType
                    });
                }
                if (SaveErrorsExcel.Count <= 3 && SaveErrorsExcel.Count >= 0)
                {
                    //Guardamos los datos en Nova        
                    ResultsInserts = GuardarListaDeNovedades(ListNovedades, ListAusetismos, ListRetiros, RuteCreate.Date, RuteCreate.Agre);

                    if (ResultsInserts == 0)
                    {
                        Alerts("Novedades, ausentismos y retiros guardados correctamente.", NotificationTypes.success);
                        hc.GuardarLogs("El cliente ha ingresado el archivo masivo de noveades de nómina",
                                                   "Proceso correcto, archivo ingresado correctamente",
                                                   User.Identity.Name,
                                                   Session["NIT"].ToString());
                        return RedirectToAction("NoveltiesAndAbsences",
                                                new ToPrincipalView
                                                {
                                                    date = RuteCreate.Date,
                                                    Agree = RuteCreate.Agre,
                                                    IncomeType = RuteCreate.IncomeType
                                                });
                    }
                    if (ResultsInserts == 1)
                    {
                        Alerts("No hemos podido guardar tus novedades, ausentismos y retiros, inténtalo más tarde.", NotificationTypes.error);
                        hc.GuardarLogs("No podemos guardar los datos de las entradas y datos de guardados del excel de novedades. El metodo de guardado por listas contesto False en algun momento del proceso", response.StatusDescription, NameCompany(), Session["NIT"].ToString());
                        return RedirectToAction("NoveltiesAndAbsences", new ToPrincipalView
                        {
                            date = RuteCreate.Date,
                            Agree = RuteCreate.Agre,
                            IncomeType = RuteCreate.IncomeType
                        });
                    }
                    Alerts("Hemos detectado algunos registros con inconsistencias que intentaste guardar. <br />Revísalas con el botón de: novedades mal ingresadas por excel.", NotificationTypes.warning);
                    hc.GuardarLogs("El cliente ha ingresado el archivo masivo de noveades de nómina",
                                               "Proceso correcto, archivo ingresado correctamente",
                                               User.Identity.Name,
                                               Session["NIT"].ToString());
                    return RedirectToAction("NoveltiesAndAbsences",
                                            new ToPrincipalView
                                            {
                                                date = RuteCreate.Date,
                                                Agree = RuteCreate.Agre,
                                                IncomeType = RuteCreate.IncomeType
                                            });
                }
                else
                {
                    hc.GuardarLogs("El cliente ha revisado el archivo masivo de noveades de nómina con errores",
                                          "Proceso correcto, archivo ingresado correctamente",
                                          User.Identity.Name,
                                          Session["NIT"].ToString());
                    Alerts("Se ha revisado el archivo ingresado y las novedades presentan errores de estructura, debes modificar tu documento o no podremos guardar tus novedades", NotificationTypes.warning);
                    return RedirectToAction("NoveltiesAndAbsences",
                                            new ToPrincipalView
                                            {
                                                date = RuteCreate.Date,
                                                Agree = RuteCreate.Agre,
                                                IncomeType = RuteCreate.IncomeType
                                            });
                }
            }
            catch (Exception ex)
            {
                hc.GuardarLogs(ex.Message, "Error, no es posible registrar un archivo excel entregado por el cliente.", NameCompany(), Session["NIT"].ToString());
                Alerts("No pudimos registrar tu archivo.", NotificationTypes.error);
                return View("Error");
            }
        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult GenerateExcelEmployees(RuteToCreate RuteCreate)
        {
            try
            {
                //Se generan las variables de session
                if (!SessionApi())
                {
                    hc.GuardarLogs("Variables de session NO generadas desde el metodo GenerateExcelEmployees en el controlador Transactions.",
                "Proceso incompleto, no se pudieron generar las variables de sesión para el proceso",
                User.Identity.Name,
                "N/A");
                    return View("ErrorApp");
                }
                //Aplico licencia 
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                //Inicio el documento de excel
                ExcelPackage Excel = new ExcelPackage();
                Excel.Workbook.Properties.Title = "ME listado de empleados";
                Excel.Workbook.Properties.Author = "Misión Empresarial";
                //Creo las hojas del archivo
                var Empleados = Excel.Workbook.Worksheets.Add("Empleados");
                //Busco la información para generar el archivo
                searchData sd = new searchData();
                sd.Date = Convert.ToDateTime(RuteCreate.Date).ToString("yyyyMMdd");
                sd.Agreement = RuteCreate.Agre;
                //Estraigo los empleados del convenio de la compañia
                IRestResponse response = new ApiUrl().UrlExecute("PaysheetNovelty/EmployeeCostStructure", Session["ApiToken"].ToString(), "POST", JsonConvert.SerializeObject(sd));
                if (response.StatusDescription != "OK")
                {
                    string descripcion = "Error al generar el listado de empleados para la generación del excel.";
                    hc.GuardarLogs(descripcion, response.StatusDescription, NameCompany(), Session["NIT"].ToString());
                    Alerts("No pudimos generar el listado de empleados, inténtalo de nuevo.", NotificationTypes.error);
                    return RedirectToAction("NoveltiesAndAbsences", new ToPrincipalView { date = RuteCreate.Date, Agree = RuteCreate.Agre, IncomeType = RuteCreate.IncomeType });
                }
                var ListEmps = JsonConvert.DeserializeObject<List<employeeCostStructure>>(response.Content);
                var ListEmp = ListEmps.OrderBy(x => x.strNombreEmpleado);
                //Se deben bloquear las cedulas y nombres de las personas
                Empleados.Protection.IsProtected = true;
                int filas = 2;
                Empleados.Cells[1, 1].Value = "CÉDULA";
                Empleados.Cells[1, 1].Style.Font.Size = 13;
                Empleados.Cells[1, 2].Value = "NOMBRE EMPLEADOS";
                Empleados.Cells[1, 2].Style.Font.Size = 13;
                Empleados.Cells[1, 3].Value = "CONVENIO";
                Empleados.Cells[1, 3].Style.Font.Size = 13;
                Empleados.Cells[1, 4].Value = "SUCURSAL";
                Empleados.Cells[1, 4].Style.Font.Size = 13;
                Empleados.Cells[1, 5].Value = "CENTRO DE COSTOS";
                Empleados.Cells[1, 5].Style.Font.Size = 13;
                Empleados.Cells[1, 6].Value = "CLASIFICADOR 1";
                Empleados.Cells[1, 6].Style.Font.Size = 13;
                Empleados.Cells[1, 7].Value = "CLASIFICADOR 2";
                Empleados.Cells[1, 7].Style.Font.Size = 13;
                Empleados.Cells[1, 8].Value = "CLASIFICADOR 3";
                Empleados.Cells[1, 8].Style.Font.Size = 13;
                Empleados.Cells[1, 9].Value = "CLASIFICADOR 4";
                Empleados.Cells[1, 9].Style.Font.Size = 13;
                Empleados.Cells[1, 10].Value = "CLASIFICADOR 5";
                Empleados.Cells[1, 10].Style.Font.Size = 13;
                Empleados.Cells[1, 11].Value = "CLASIFICADOR 6";
                Empleados.Cells[1, 11].Style.Font.Size = 13;
                Empleados.Cells[1, 12].Value = "CLASIFICADOR 7";
                Empleados.Cells[1, 12].Style.Font.Size = 13;
                Empleados.Cells[1, 13].Value = "CLASIFICADOR 8";
                Empleados.Cells[1, 13].Style.Font.Size = 13;
                Empleados.Cells[1, 14].Value = "FECHA DE INGRESO";
                Empleados.Cells[1, 14].Style.Font.Size = 13;
                foreach (var item in ListEmp)
                {
                    Empleados.Cells[filas, 1].Value = item.strCodigoEmpleado;
                    Empleados.Cells[filas, 2].Value = item.strNombreEmpleado;
                    Empleados.Cells[filas, 3].Value = item.strConvenio;
                    Empleados.Cells[filas, 4].Value = item.strSucursal;
                    Empleados.Cells[filas, 5].Value = item.strCentroDeCostos;
                    Empleados.Cells[filas, 6].Value = item.strClasificador1;
                    Empleados.Cells[filas, 7].Value = item.strClasificador2;
                    Empleados.Cells[filas, 8].Value = item.strClasificador3;
                    Empleados.Cells[filas, 9].Value = item.strClasificador4;
                    Empleados.Cells[filas, 10].Value = item.strClasificador5;
                    Empleados.Cells[filas, 11].Value = item.strClasificador6;
                    Empleados.Cells[filas, 12].Value = item.strClasificador7;
                    Empleados.Cells[filas, 13].Value = item.strClasificador8;
                    Empleados.Cells[filas, 14].Value = item.strFechaIngreso;
                    filas++;
                }
                Empleados.Row(1).Height = 41.25;
                Empleados.Cells[1, 1, 1, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                Empleados.Cells[1, 1, 1, 14].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F4F6F8"));
                Empleados.Cells[1, 1, 1, 14].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                Empleados.Cells[1, 1, 1, 14].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                Empleados.Cells[1, 1, 1, 14].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                Empleados.Cells[1, 1, 1, 14].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                Empleados.Cells[1, 1, 1, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Empleados.Cells[1, 1, 1, 14].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                Empleados.Cells[1, 1, 1, 14].Style.Font.Bold = true;
                Empleados.Cells[Empleados.Dimension.Address].Style.Font.Name = "Century Gothic";
                Empleados.Cells[Empleados.Dimension.Address].AutoFitColumns();
                //Hoja de Ausentismos               
                hc.GuardarLogs("El cliente genero el archivo con los empleados activos correctamente",
                                           "Proceso correcto, el cliente ha generado el archivo de empleados activos",
                                           User.Identity.Name,
                                           Session["NIT"].ToString());
                return File(Excel.GetAsByteArray(), "application/octet-stream", "Empleados activos a " + sd.Date + ".xlsx");
            }
            catch (Exception ex)
            {
                string descripcion = "Error al generar en la generación del excel.";
                hc.GuardarLogs(descripcion, ex.Message, NameCompany(), Session["NIT"].ToString());
                Alerts("No pudimos generar tu archivo.", NotificationTypes.error);
                return RedirectToAction("NoveltiesAndAbsences", new ToPrincipalView { date = RuteCreate.Date, Agree = RuteCreate.Agre, IncomeType = RuteCreate.IncomeType });
            }
        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Prenomina_Register(RuteToCreate RuteCreate)
        {
            try
            {
                //Se generan las variables de session
                if (!SessionApi())
                {
                    hc.GuardarLogs("Variables de session NO generadas desde el metodo Prenomina_Register en el controlador Transactions.",
                                   "Proceso incompleto, no se pudieron generar las variables de sesión para el proceso",
                                   User.Identity.Name,
                                   "N/A");
                    return View("ErrorApp");
                }
                //Aplico licencia 
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                //Inicio el documento de excel
                ExcelPackage Excel = new ExcelPackage();
                Excel.Workbook.Properties.Title = "ME novedades ingresadas para " + RuteCreate.Date;
                Excel.Workbook.Properties.Author = "Misión Empresarial";
                //Creo las hojas del archivo
                var Novedades = Excel.Workbook.Worksheets.Add("Novedades");
                var Ausentismos = Excel.Workbook.Worksheets.Add("Ausentismos");
                var Retiros = Excel.Workbook.Worksheets.Add("Retiros");
                //Busco la información para generar el archivo
                QueryByCustomer qc = new QueryByCustomer();
                qc.Date = RuteCreate.Date;
                qc.Nit = Session["NIT"].ToString();
                qc.Agreement = RuteCreate.Agre;
                IRestResponse response = new ApiUrl().UrlExecute("PaysheetNovelty/Customer", Session["ApiToken"].ToString(), "POST", JsonConvert.SerializeObject(qc));
                if (response.StatusDescription != "OK")
                {
                    hc.GuardarLogs("Error al traer la información de novedades y ausentismos.", response.StatusDescription, NameCompany(), Session["NIT"].ToString());
                    Alerts("No pudimos cargar la información de novedades y ausentismos, inténtalo de nuevo.", NotificationTypes.error);
                    return RedirectToAction("NoveltiesAndAbsences", new ToPrincipalView { date = RuteCreate.Date, Agree = RuteCreate.Agre, IncomeType = RuteCreate.IncomeType });
                }
                List<PaysheetNoveltyData> ListNovelties = JsonConvert.DeserializeObject<List<PaysheetNoveltyData>>(response.Content);
                searchData sd = new searchData();
                sd.Date = Convert.ToDateTime(RuteCreate.Date).ToString("yyyyMMdd");
                sd.Agreement = RuteCreate.Agre;
                response = new ApiUrl().UrlExecute("PaysheetNovelty/EmployeeCostStructure", Session["ApiToken"].ToString(), "POST", JsonConvert.SerializeObject(sd));
                if (response.StatusDescription != "OK")
                {
                    string descripcion = "Error al generar el listado de empleados para la generación del excel.";
                    hc.GuardarLogs(descripcion, response.StatusDescription, NameCompany(), Session["NIT"].ToString());
                    Alerts("No pudimos generar el listado de empleados, inténtalo de nuevo.", NotificationTypes.error);
                    return RedirectToAction("NoveltiesAndAbsences", new ToPrincipalView { date = RuteCreate.Date, Agree = RuteCreate.Agre, IncomeType = RuteCreate.IncomeType });
                }
                List<employeeCostStructure> ListEmps = JsonConvert.DeserializeObject<List<employeeCostStructure>>(response.Content);
                //Comienzo llenado la lista de novedades y retiros.
                //Se deben bloquear las cedulas y nombres de las personas
                Novedades.Protection.IsProtected = true;
                Novedades.Cells[1, 1].Value = "CÉDULA";
                Novedades.Cells[1, 1].Style.Font.Size = 13;
                Novedades.Cells[1, 2].Value = "NOMBRE EMPLEADOS";
                Novedades.Cells[1, 2].Style.Font.Size = 13;
                Novedades.Cells[1, 3].Value = "CONVENIO";
                Novedades.Cells[1, 3].Style.Font.Size = 13;
                Novedades.Cells[1, 4].Value = "SUCURSAL";
                Novedades.Cells[1, 4].Style.Font.Size = 13;
                Novedades.Cells[1, 5].Value = "CENTRO DE COSTOS";
                Novedades.Cells[1, 5].Style.Font.Size = 13;
                Novedades.Cells[1, 6].Value = "CLASIFICADOR 1";
                Novedades.Cells[1, 6].Style.Font.Size = 13;
                Novedades.Cells[1, 7].Value = "CLASIFICADOR 2";
                Novedades.Cells[1, 7].Style.Font.Size = 13;
                Novedades.Cells[1, 8].Value = "CLASIFICADOR 3";
                Novedades.Cells[1, 8].Style.Font.Size = 13;
                Novedades.Cells[1, 9].Value = "CLASIFICADOR 4";
                Novedades.Cells[1, 9].Style.Font.Size = 13;
                Novedades.Cells[1, 10].Value = "CLASIFICADOR 5";
                Novedades.Cells[1, 10].Style.Font.Size = 13;
                Novedades.Cells[1, 11].Value = "CLASIFICADOR 6";
                Novedades.Cells[1, 11].Style.Font.Size = 13;
                Novedades.Cells[1, 12].Value = "CLASIFICADOR 7";
                Novedades.Cells[1, 12].Style.Font.Size = 13;
                Novedades.Cells[1, 13].Value = "CLASIFICADOR 8";
                Novedades.Cells[1, 13].Style.Font.Size = 13;
                Novedades.Cells[1, 14].Value = "TIPO CONCEPTO";
                Novedades.Cells[1, 14].Style.Font.Size = 13;
                Novedades.Cells[1, 15].Value = "VALOR";
                Novedades.Cells[1, 15].Style.Font.Size = 13;
                Novedades.Cells[1, 1, 1, 15].Style.Fill.PatternType = ExcelFillStyle.Solid;
                Novedades.Cells[1, 1, 1, 15].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F4F6F8"));
                Novedades.Cells[1, 1, 1, 15].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                Novedades.Cells[1, 1, 1, 15].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                Novedades.Cells[1, 1, 1, 15].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                Novedades.Cells[1, 1, 1, 15].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                Novedades.Cells[1, 1, 1, 15].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Novedades.Cells[1, 1, 1, 15].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                Novedades.Cells[1, 1, 1, 15].Style.Font.Bold = true;
                Novedades.Cells[Novedades.Dimension.Address].Style.Font.Name = "Century Gothic";
                //Hoja de Ausentismos
                Ausentismos.Protection.IsProtected = true;
                Ausentismos.Cells[1, 1].Value = "CÉDULA";
                Ausentismos.Cells[1, 1].Style.Font.Size = 13;
                Ausentismos.Cells[1, 2].Value = "NOMBRE EMPLEADOS";
                Ausentismos.Cells[1, 2].Style.Font.Size = 13;
                Ausentismos.Cells[1, 3].Value = "CONVENIO";
                Ausentismos.Cells[1, 3].Style.Font.Size = 13;
                Ausentismos.Cells[1, 4].Value = "SUCURSAL";
                Ausentismos.Cells[1, 4].Style.Font.Size = 13;
                Ausentismos.Cells[1, 5].Value = "CENTRO DE COSTOS";
                Ausentismos.Cells[1, 5].Style.Font.Size = 13;
                Ausentismos.Cells[1, 6].Value = "CLASIFICADOR 1";
                Ausentismos.Cells[1, 6].Style.Font.Size = 13;
                Ausentismos.Cells[1, 7].Value = "CLASIFICADOR 2";
                Ausentismos.Cells[1, 7].Style.Font.Size = 13;
                Ausentismos.Cells[1, 8].Value = "CLASIFICADOR 3";
                Ausentismos.Cells[1, 8].Style.Font.Size = 13;
                Ausentismos.Cells[1, 9].Value = "CLASIFICADOR 4";
                Ausentismos.Cells[1, 9].Style.Font.Size = 13;
                Ausentismos.Cells[1, 10].Value = "CLASIFICADOR 5";
                Ausentismos.Cells[1, 10].Style.Font.Size = 13;
                Ausentismos.Cells[1, 11].Value = "CLASIFICADOR 6";
                Ausentismos.Cells[1, 11].Style.Font.Size = 13;
                Ausentismos.Cells[1, 12].Value = "CLASIFICADOR 7";
                Ausentismos.Cells[1, 12].Style.Font.Size = 13;
                Ausentismos.Cells[1, 13].Value = "CLASIFICADOR 8";
                Ausentismos.Cells[1, 13].Style.Font.Size = 13;
                Ausentismos.Cells[1, 14].Value = "TIPO AUSENTISMO";
                Ausentismos.Cells[1, 14].Style.Font.Size = 13;
                Ausentismos.Cells[1, 15].Value = "FECHA INICIO";
                Ausentismos.Cells[1, 15].Style.Font.Size = 13;
                Ausentismos.Cells[1, 16].Value = "FECHA FINAL";
                Ausentismos.Cells[1, 16].Style.Font.Size = 13;
                Ausentismos.Cells[1, 1, 1, 16].Style.Fill.PatternType = ExcelFillStyle.Solid;
                Ausentismos.Cells[1, 1, 1, 16].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F4F6F8"));
                Ausentismos.Cells[1, 1, 1, 16].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                Ausentismos.Cells[1, 1, 1, 16].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                Ausentismos.Cells[1, 1, 1, 16].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                Ausentismos.Cells[1, 1, 1, 16].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                Ausentismos.Cells[1, 1, 1, 16].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Ausentismos.Cells[1, 1, 1, 16].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                Ausentismos.Cells[1, 1, 1, 16].Style.Font.Bold = true;
                Ausentismos.Cells[Ausentismos.Dimension.Address].Style.Font.Name = "Century Gothic";
                //Hoja de retirtos
                Retiros.Protection.IsProtected = true;
                Retiros.Cells[1, 1].Value = "CÉDULA";
                Retiros.Cells[1, 1].Style.Font.Size = 13;
                Retiros.Cells[1, 2].Value = "NOMBRE EMPLEADOS";
                Retiros.Cells[1, 2].Style.Font.Size = 13;
                Retiros.Cells[1, 3].Value = "CONVENIO";
                Retiros.Cells[1, 3].Style.Font.Size = 13;
                Retiros.Cells[1, 4].Value = "SUCURSAL";
                Retiros.Cells[1, 4].Style.Font.Size = 13;
                Retiros.Cells[1, 5].Value = "CENTRO DE COSTOS";
                Retiros.Cells[1, 5].Style.Font.Size = 13;
                Retiros.Cells[1, 6].Value = "CLASIFICADOR 1";
                Retiros.Cells[1, 6].Style.Font.Size = 13;
                Retiros.Cells[1, 7].Value = "CLASIFICADOR 2";
                Retiros.Cells[1, 7].Style.Font.Size = 13;
                Retiros.Cells[1, 8].Value = "CLASIFICADOR 3";
                Retiros.Cells[1, 8].Style.Font.Size = 13;
                Retiros.Cells[1, 9].Value = "CLASIFICADOR 4";
                Retiros.Cells[1, 9].Style.Font.Size = 13;
                Retiros.Cells[1, 10].Value = "CLASIFICADOR 5";
                Retiros.Cells[1, 10].Style.Font.Size = 13;
                Retiros.Cells[1, 11].Value = "CLASIFICADOR 6";
                Retiros.Cells[1, 11].Style.Font.Size = 13;
                Retiros.Cells[1, 12].Value = "CLASIFICADOR 7";
                Retiros.Cells[1, 12].Style.Font.Size = 13;
                Retiros.Cells[1, 13].Value = "CLASIFICADOR 8";
                Retiros.Cells[1, 13].Style.Font.Size = 13;
                Retiros.Cells[1, 14].Value = "FECHA DE RETIRO";
                Retiros.Cells[1, 14].Style.Font.Size = 13;
                Retiros.Cells[1, 1, 1, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                Retiros.Cells[1, 1, 1, 14].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F4F6F8"));
                Retiros.Cells[1, 1, 1, 14].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                Retiros.Cells[1, 1, 1, 14].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                Retiros.Cells[1, 1, 1, 14].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                Retiros.Cells[1, 1, 1, 14].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                Retiros.Cells[1, 1, 1, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Retiros.Cells[1, 1, 1, 14].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                Retiros.Cells[1, 1, 1, 14].Style.Font.Bold = true;
                Retiros.Cells[Retiros.Dimension.Address].Style.Font.Name = "Century Gothic";
                List<PaysheetNovelty> ListNov = new List<PaysheetNovelty>();
                List<PaysheetNovelty> ListAbs = new List<PaysheetNovelty>();
                List<PaysheetNovelty> ListRet = new List<PaysheetNovelty>();
                int Filas = 2;
                foreach (var item in ListNovelties)
                {
                    foreach (var Novedad in item.Novelty)
                    {
                        ListNov.Add(Novedad);
                    }
                    foreach (var Ausentismo in item.Absenteeism)
                    {
                        ListAbs.Add(Ausentismo);
                    }
                    foreach (var Retiro in item.Retirement)
                    {
                        ListRet.Add(Retiro);
                    }
                }
                foreach (var Novedad in ListNov)
                {
                    var Emplo = ListEmps.Select(x => x).Where(x => x.strCodigoEmpleado.Trim() == Novedad.Employee_Id.Trim()).FirstOrDefault();
                    if (Emplo != null)
                    {
                        Novedades.Cells[Filas, 1].Value = Emplo.strCodigoEmpleado;
                        Novedades.Cells[Filas, 2].Value = Emplo.strNombreEmpleado;
                        Novedades.Cells[Filas, 3].Value = Emplo.strConvenio;
                        Novedades.Cells[Filas, 4].Value = Emplo.strSucursal;
                        Novedades.Cells[Filas, 5].Value = Emplo.strCentroDeCostos;
                        Novedades.Cells[Filas, 6].Value = Emplo.strClasificador1;
                        Novedades.Cells[Filas, 7].Value = Emplo.strClasificador2;
                        Novedades.Cells[Filas, 8].Value = Emplo.strClasificador3;
                        Novedades.Cells[Filas, 9].Value = Emplo.strClasificador4;
                        Novedades.Cells[Filas, 10].Value = Emplo.strClasificador5;
                        Novedades.Cells[Filas, 11].Value = Emplo.strClasificador6;
                        Novedades.Cells[Filas, 12].Value = Emplo.strClasificador7;
                        Novedades.Cells[Filas, 13].Value = Emplo.strClasificador8;
                    }
                    Novedades.Cells[Filas, 14].Value = Novedad.Concept_Name;
                    Novedades.Cells[Filas, 15].Value = Novedad.Value;
                    Filas++;
                }
                Filas = 2;
                foreach (var Ausentismo in ListAbs)
                {
                    var Emplo = ListEmps.Select(x => x).Where(x => x.strCodigoEmpleado.Trim() == Ausentismo.Employee_Id.Trim()).FirstOrDefault();
                    if (Emplo != null)
                    {
                        Ausentismos.Cells[Filas, 1].Value = Emplo.strCodigoEmpleado;
                        Ausentismos.Cells[Filas, 2].Value = Emplo.strNombreEmpleado;
                        Ausentismos.Cells[Filas, 3].Value = Emplo.strConvenio;
                        Ausentismos.Cells[Filas, 4].Value = Emplo.strSucursal;
                        Ausentismos.Cells[Filas, 5].Value = Emplo.strCentroDeCostos;
                        Ausentismos.Cells[Filas, 6].Value = Emplo.strClasificador1;
                        Ausentismos.Cells[Filas, 7].Value = Emplo.strClasificador2;
                        Ausentismos.Cells[Filas, 8].Value = Emplo.strClasificador3;
                        Ausentismos.Cells[Filas, 9].Value = Emplo.strClasificador4;
                        Ausentismos.Cells[Filas, 10].Value = Emplo.strClasificador5;
                        Ausentismos.Cells[Filas, 11].Value = Emplo.strClasificador6;
                        Ausentismos.Cells[Filas, 12].Value = Emplo.strClasificador7;
                        Ausentismos.Cells[Filas, 13].Value = Emplo.strClasificador8;
                    }
                    Ausentismos.Cells[Filas, 14].Value = Ausentismo.Absenteeism_Name;
                    Ausentismos.Cells[Filas, 15].Value = Ausentismo.DateStart;
                    Ausentismos.Cells[Filas, 16].Value = Ausentismo.DateEnd;
                    Filas++;
                }
                Filas = 2;
                foreach (var Retiro in ListRet)
                {
                    var Emplo = ListEmps.Select(x => x).Where(x => x.strCodigoEmpleado.Trim() == Retiro.Employee_Id.Trim()).FirstOrDefault();
                    if (Emplo != null)
                    {
                        Retiros.Cells[Filas, 1].Value = Emplo.strCodigoEmpleado;
                        Retiros.Cells[Filas, 2].Value = Emplo.strNombreEmpleado;
                        Retiros.Cells[Filas, 3].Value = Emplo.strConvenio;
                        Retiros.Cells[Filas, 4].Value = Emplo.strSucursal;
                        Retiros.Cells[Filas, 5].Value = Emplo.strCentroDeCostos;
                        Retiros.Cells[Filas, 6].Value = Emplo.strClasificador1;
                        Retiros.Cells[Filas, 7].Value = Emplo.strClasificador2;
                        Retiros.Cells[Filas, 8].Value = Emplo.strClasificador3;
                        Retiros.Cells[Filas, 9].Value = Emplo.strClasificador4;
                        Retiros.Cells[Filas, 10].Value = Emplo.strClasificador5;
                        Retiros.Cells[Filas, 11].Value = Emplo.strClasificador6;
                        Retiros.Cells[Filas, 12].Value = Emplo.strClasificador7;
                        Retiros.Cells[Filas, 13].Value = Emplo.strClasificador8;
                    }
                    Retiros.Cells[Filas, 14].Value = Retiro.DateEnd;
                    Filas++;
                }
                Retiros.Cells[Retiros.Dimension.Address].AutoFitColumns();
                Novedades.Cells[Novedades.Dimension.Address].AutoFitColumns();
                Ausentismos.Cells[Ausentismos.Dimension.Address].AutoFitColumns();
                hc.GuardarLogs("El cliente genera el archivo de los registros actuales de la nómina",
                                           "Proceso correcto, generación de registros de nómina actual",
                                           User.Identity.Name,
                                           Session["NIT"].ToString());
                return File(Excel.GetAsByteArray(), "application/octet-stream", "Novedades ingresadas " + RuteCreate.Agre + "_" + RuteCreate.Date + ".xlsx");
            }
            catch (Exception ex)
            {
                hc.GuardarLogs("Error en la generación del excel de novedades registradas.", ex.Message, NameCompany(), Session["NIT"].ToString());
                Alerts("No pudimos generar tu archivo", NotificationTypes.error);
                return RedirectToAction("NoveltiesAndAbsences", new ToPrincipalView { date = RuteCreate.Date, Agree = RuteCreate.Agre, IncomeType = RuteCreate.IncomeType });
            }
        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult FileListErrorSaveEntryExcel(RuteToCreate RuteCreate)
        {
            try
            {
                //Se generan las variables de session
                if (!SessionApi())
                {
                    hc.GuardarLogs("Variables de session NO generadas desde el metodo FileListErrorSaveEntryExcel en el controlador Transactions.",
               "Proceso incompleto, no se pudieron generar las variables de sesión para el proceso",
               User.Identity.Name,
               "N/A");
                    return View("ErrorApp");
                }
                List<PaysheetEntryAndSaveErrorsExcel> SaveErrorsExcel = new List<PaysheetEntryAndSaveErrorsExcel>();
                NoveltyData nd = new NoveltyData();
                nd.Agreement = RuteCreate.Agre;
                nd.Date = RuteCreate.Date;
                IRestResponse response = new ApiUrl().UrlExecute("PaysheetEntryAndSaveErrorsExcel/ListPaysheetEntryErrors", Session["ApiToken"].ToString(), "POST", JsonConvert.SerializeObject(nd));
                if (response.StatusDescription != "OK")
                {
                    hc.GuardarLogs("Error al traer la información de las novedades de ingreso de masio por excel.", response.StatusDescription, NameCompany(), Session["NIT"].ToString());
                    Alerts("No pudimos cargar la información del archivo cargado por la plataforma, inténtalo de nuevo.", NotificationTypes.error);
                    return RedirectToAction("NoveltiesAndAbsences", new ToPrincipalView { date = RuteCreate.Date, Agree = RuteCreate.Agre, IncomeType = RuteCreate.IncomeType });
                }
                SaveErrorsExcel = JsonConvert.DeserializeObject<List<PaysheetEntryAndSaveErrorsExcel>>(response.Content);
                //Aplico licencia 
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                //Inicio el documento de excel
                ExcelPackage Excel = new ExcelPackage();
                Excel.Workbook.Properties.Title = "ME errores de ingreso y guardado por excel masivo";
                Excel.Workbook.Properties.Author = "Misión Empresarial";
                //Creo las hojas del archivo
                var BadEntry = Excel.Workbook.Worksheets.Add("Malos ingresos");
                BadEntry.Protection.IsProtected = true;
                BadEntry.Cells[1, 1].Value = "EMPLEADO";
                BadEntry.Cells[1, 1].Style.Font.Size = 13;
                BadEntry.Cells[1, 2].Value = "CÓDIGO CONCEPTO";
                BadEntry.Cells[1, 2].Style.Font.Size = 13;
                BadEntry.Cells[1, 3].Value = "NOMBRE CONCEPTO";
                BadEntry.Cells[1, 3].Style.Font.Size = 13;
                BadEntry.Cells[1, 4].Value = "VALOR";
                BadEntry.Cells[1, 4].Style.Font.Size = 13;
                BadEntry.Cells[1, 5].Value = "CÓDIOGO AUSENTISMO";
                BadEntry.Cells[1, 5].Style.Font.Size = 13;
                BadEntry.Cells[1, 6].Value = "NOMBRE AUSENTISMO";
                BadEntry.Cells[1, 6].Style.Font.Size = 13;
                BadEntry.Cells[1, 7].Value = "FECHA INICIAL REGISTRO AUSENTISMO";
                BadEntry.Cells[1, 7].Style.Font.Size = 13;
                BadEntry.Cells[1, 8].Value = "FECHA FINAL AUSENTISMO/FECHA DE RETIRO";
                BadEntry.Cells[1, 8].Style.Font.Size = 13;
                BadEntry.Cells[1, 9].Value = "FECHA REAL INICIAL AUSENTISMO";
                BadEntry.Cells[1, 9].Style.Font.Size = 13;
                BadEntry.Cells[1, 10].Value = "FECHA REAL FINAL AUSENTISMO";
                BadEntry.Cells[1, 10].Style.Font.Size = 13;
                BadEntry.Cells[1, 11].Value = "Error";
                BadEntry.Cells[1, 11].Style.Font.Size = 13;
                BadEntry.Cells[1, 1, 1, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                BadEntry.Cells[1, 1, 1, 11].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F4F6F8"));
                BadEntry.Cells[1, 1, 1, 11].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                BadEntry.Cells[1, 1, 1, 11].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                BadEntry.Cells[1, 1, 1, 11].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                BadEntry.Cells[1, 1, 1, 11].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                BadEntry.Cells[1, 1, 1, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                BadEntry.Cells[1, 1, 1, 11].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                BadEntry.Cells[1, 1, 1, 11].Style.Font.Bold = true;
                BadEntry.Cells[BadEntry.Dimension.Address].Style.Font.Name = "Century Gothic";
                var BadSave = Excel.Workbook.Worksheets.Add("No guardados");
                BadSave.Protection.IsProtected = true;
                BadSave.Cells[1, 1].Value = "EMPLEADO";
                BadSave.Cells[1, 1].Style.Font.Size = 13;
                BadSave.Cells[1, 2].Value = "CÓDIGO CONCEPTO";
                BadSave.Cells[1, 2].Style.Font.Size = 13;
                BadSave.Cells[1, 3].Value = "NOMBRE CONCEPTO";
                BadSave.Cells[1, 3].Style.Font.Size = 13;
                BadSave.Cells[1, 4].Value = "VALOR";
                BadSave.Cells[1, 4].Style.Font.Size = 13;
                BadSave.Cells[1, 5].Value = "CÓDIOGO AUSENTISMO";
                BadSave.Cells[1, 5].Style.Font.Size = 13;
                BadSave.Cells[1, 6].Value = "NOMBRE AUSENTISMO";
                BadSave.Cells[1, 6].Style.Font.Size = 13;
                BadSave.Cells[1, 7].Value = "FECHA INICIAL REGISTRO AUSENTISMO";
                BadSave.Cells[1, 7].Style.Font.Size = 13;
                BadSave.Cells[1, 8].Value = "FECHA FINAL AUSENTISMO/FECHA DE RETIRO";
                BadSave.Cells[1, 8].Style.Font.Size = 13;
                BadSave.Cells[1, 9].Value = "FECHA REAL INICIAL AUSENTISMO";
                BadSave.Cells[1, 9].Style.Font.Size = 13;
                BadSave.Cells[1, 10].Value = "FECHA REAL FINAL AUSENTISMO";
                BadSave.Cells[1, 10].Style.Font.Size = 13;
                BadSave.Cells[1, 11].Value = "MENSAJE DE ERROR";
                BadSave.Cells[1, 11].Style.Font.Size = 13;
                BadSave.Cells[1, 1, 1, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                BadSave.Cells[1, 1, 1, 11].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F4F6F8"));
                BadSave.Cells[1, 1, 1, 11].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                BadSave.Cells[1, 1, 1, 11].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                BadSave.Cells[1, 1, 1, 11].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                BadSave.Cells[1, 1, 1, 11].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                BadSave.Cells[1, 1, 1, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                BadSave.Cells[1, 1, 1, 11].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                BadSave.Cells[1, 1, 1, 11].Style.Font.Bold = true;
                BadSave.Cells[BadSave.Dimension.Address].Style.Font.Name = "Century Gothic";
                int Filas = 2;
                var ListBadEntry = SaveErrorsExcel.Where(x => x.Flag == 1 || x.Flag == 2 || x.Flag == 3).Select(x => x).ToList();
                foreach (var item in ListBadEntry)
                {
                    if (item.Flag == 1)
                    {
                        BadEntry.Cells[Filas, 1].Value = item.Employee;
                        BadEntry.Cells[Filas, 2].Value = item.Concept_Id;
                        BadEntry.Cells[Filas, 3].Value = item.Concept_Name;
                        BadEntry.Cells[Filas, 4].Value = item.Values;
                        BadEntry.Cells[Filas, 11].Value = item.MessageError;
                        Filas++;
                    }
                    else
                    {
                        if (item.Flag == 2)
                        {
                            BadEntry.Cells[Filas, 1].Value = item.Employee;
                            BadEntry.Cells[Filas, 5].Value = item.Absenteeism_Id;
                            BadEntry.Cells[Filas, 6].Value = item.Absenteeism_Name;
                            BadEntry.Cells[Filas, 7].Value = item.DateStart;
                            BadEntry.Cells[Filas, 8].Value = item.DateEnd;
                            BadEntry.Cells[Filas, 9].Value = item.DateReallyStart;
                            BadEntry.Cells[Filas, 10].Value = item.DateReallyEnd;
                            BadEntry.Cells[Filas, 11].Value = item.MessageError;
                            Filas++;
                        }
                        else
                        {
                            if (item.Flag == 3)
                            {
                                BadEntry.Cells[Filas, 1].Value = item.Employee;
                                BadEntry.Cells[Filas, 8].Value = item.DateEnd;
                                BadEntry.Cells[Filas, 11].Value = item.MessageError;
                                Filas++;
                            }
                        }
                    }
                }
                Filas = 2;
                var ListBadSave = SaveErrorsExcel.Where(x => x.Flag == 4 || x.Flag == 5 || x.Flag == 6).Select(x => x).ToList();
                foreach (var item in ListBadSave)
                {
                    if (item.Flag == 4)
                    {
                        BadSave.Cells[Filas, 1].Value = item.Employee;
                        BadSave.Cells[Filas, 2].Value = item.Concept_Id;
                        BadSave.Cells[Filas, 3].Value = item.Concept_Name;
                        BadSave.Cells[Filas, 4].Value = item.Values;
                        BadSave.Cells[Filas, 11].Value = item.MessageError;
                        Filas++;
                    }
                    else
                    {
                        if (item.Flag == 5)
                        {
                            BadSave.Cells[Filas, 1].Value = item.Employee;
                            BadSave.Cells[Filas, 5].Value = item.Absenteeism_Id;
                            BadSave.Cells[Filas, 6].Value = item.Absenteeism_Name;
                            BadSave.Cells[Filas, 7].Value = item.DateStart;
                            BadSave.Cells[Filas, 8].Value = item.DateEnd;
                            BadSave.Cells[Filas, 9].Value = item.DateReallyStart;
                            BadSave.Cells[Filas, 10].Value = item.DateReallyEnd;
                            BadSave.Cells[Filas, 11].Value = item.MessageError;
                            Filas++;
                        }
                        else
                        {
                            if (item.Flag == 6)
                            {
                                BadSave.Cells[Filas, 1].Value = item.Employee;
                                BadSave.Cells[Filas, 8].Value = item.DateEnd;
                                BadSave.Cells[Filas, 11].Value = item.MessageError;
                                Filas++;
                            }
                        }
                    }
                }
                BadEntry.Cells[BadEntry.Dimension.Address].AutoFitColumns();
                BadSave.Cells[BadSave.Dimension.Address].AutoFitColumns();
                hc.GuardarLogs("El cliente genero el archivo de errores de ingreso masivo por excel",
                               "Proceso correcto, generación de archivo de ingreso masivo",
                               User.Identity.Name,
                               Session["NIT"].ToString());
                return File(Excel.GetAsByteArray(), "application/octet-stream", "Errores Ingreso y guardado por excel masivo" + RuteCreate.Agre + "_" + DateTime.Now.ToString("ddMMyyyHHmmss") + ".xlsx");
            }
            catch (Exception ex)
            {
                hc.GuardarLogs("Error en la generación del excel de errores de ingreso y guardado de novedades de nómina.", ex.Message, NameCompany(), Session["NIT"].ToString());
                Alerts("No pudimos generar tu archivo", NotificationTypes.error);
                return RedirectToAction("NoveltiesAndAbsences", new ToPrincipalView
                {
                    date = RuteCreate.Date,
                    Agree = RuteCreate.Agre,
                    IncomeType = RuteCreate.IncomeType
                });
            }
        }
        #endregion
