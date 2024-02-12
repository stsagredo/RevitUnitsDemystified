// in case you don't want to copy and paste thw two previous blocks:
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using System.Xml;


/*
 * Tested on Revit 2023.
 * 
 * Copyright Sebastian Torres Sagredo, 2023. Under MIT Licence:
 * 
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), 
 * to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, 
 * and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
 * IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 * 
 * [Transaction attribute]:
 * This is required by the API.
 * TL; DR: Manual commands imply you create transactions and manage them.
 * ReadOnly means you can access and read the model, but not edit it, else it'll throw an Exception.
 * Although we won't be creating a transaction, it won't run the UnitsExtractor if this Attribute is not set.
 */


namespace RevitUnitsDemystified
{
    [Transaction(TransactionMode.Manual)]
    public class UnitsExtractor : IExternalCommand
    {
        // This is the entry point of the extension. Revit will starte executing code here when you press the External Add-In button.
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // We will only get the current UI document, although you can get several different objects like the UIApplication from commandData.
            Document doc = commandData.Application.ActiveUIDocument.Document;

            // For this example, we will serialise the JProperty and save it as an indented string.
            string result = JsonConvert.SerializeObject(GetSpecDataByDiscipline(DisciplineTypeId.Common, doc), Newtonsoft.Json.Formatting.Indented);

            // Save the file to the computer.
            File.WriteAllText("result.json", result);

            Debug.Print(result);

            // Done!
            return Result.Succeeded;
        }

        public static JProperty GetSpecDataByDiscipline(ForgeTypeId discipline, Document doc)
        {
            // Get the label for the input discipline.
            string disciplineLabel = LabelUtils.GetLabelForDiscipline(discipline);

            // Get a list with specs filtered by discipline.
            IList<ForgeTypeId> GetSpecs()
            {
                return (from spec in UnitUtils.GetAllMeasurableSpecs()
                        let specDiscipline = UnitUtils.GetDiscipline(spec)
                        where specDiscipline.Equals(discipline)
                        select spec).ToList();
            }

            // Create a new JArray to store each spec's serialised data.
            JArray specData = new JArray();

            // Iterate through each measureable spec inside the chosen discipline.
            foreach (ForgeTypeId spec in GetSpecs())
            {
                try
                {
                    // Get labels for this spec.
                    string specLabel = LabelUtils.GetLabelForSpec(spec);

                    // Get the set unit on this spec and its label.
                    ForgeTypeId unit = doc.GetUnits().GetFormatOptions(spec).GetUnitTypeId();
                    string unitLabel = LabelUtils.GetLabelForUnit(unit);

                    // Get the symbol for the set unit on this spec.
                    ForgeTypeId symbol = doc.GetUnits().GetFormatOptions(spec).GetSymbolTypeId();

                    // To get symbol labels, we need a special method.
                    // If the symbol label is empty, or not supported by the symbol, it will throw an Exception and fail the transaction.
                    // Therefore, it must be handled different.
                    string symbolLabel(ForgeTypeId symbolTypeId)
                    {
                        try
                        {
                            return LabelUtils.GetLabelForSymbol(symbolTypeId);
                        }
                        catch
                        {
                            return String.Empty;
                        }
                    }

                    // To get a list of all the valid symbols for a unit, iterate through FormatOptions.GetValidSymbols(unitTypeId)
                    // Then get the symbol label for each.
                    JArray validSymbols(ForgeTypeId unitTypeId)
                    {
                        JArray symbols = new JArray();
                        foreach (var item in FormatOptions.GetValidSymbols(unitTypeId))
                        {
                            JObject symbolData = new JObject()
                            {
                                ["SymbolTypeId"] = item.TypeId,
                                ["SymbolLabel"] = symbolLabel(item)
                            };
                            symbols.Add(symbolData);
                        }
                        return symbols;
                    }

                    // Create the main bundle of data for each spec.
                    JObject result = new JObject
                    {
                        // Start with the label for the spec.
                        [specLabel] = new JObject
                        {
                            // Attach the TypeId schema identifier.
                            ["SpecTypeId"] = spec.TypeId,
                            // Then add the unit attached to this spec:
                            [unitLabel] = new JObject
                            {
                                // Attach the TypeId schema identifier.
                                ["UnitTypeId"] = unit.TypeId,

                                // DefaultOptions represent the default state of the unit, without any formatting applied. 
                                ["DefaultOptions"] = new JObject
                                {
                                    ["CanHaveSymbol"] = FormatOptions.CanHaveSymbol(unit),
                                    ["CanSuppressSpaces"] = FormatOptions.CanSuppressSpaces(unit),
                                    ["CanSuppressLeadingZeros"] = FormatOptions.CanSuppressLeadingZeros(unit),
                                    ["CanSuppressTrailingZeros"] = FormatOptions.CanSuppressTrailingZeros(unit),
                                    ["CanUsePlusPrefix"] = FormatOptions.CanUsePlusPrefix(unit),
                                },

                                // SetFormatOptions represent the current formatting of the unit inside the document.
                                ["SetFormatOptions"] = new JObject
                                {
                                    ["RoundingMethod"] = doc.GetUnits().GetFormatOptions(spec).RoundingMethod.ToString(),
                                    ["UseDigitGrouping"] = doc.GetUnits().GetFormatOptions(spec).UseDigitGrouping,
                                    ["UseDefaultFormatting"] = doc.GetUnits().GetFormatOptions(spec).UseDefault,
                                    ["AreFormatOptionsValidForSpec"] = doc.GetUnits().GetFormatOptions(spec).IsValidForSpec(spec),
                                    ["HasSymbol"] = doc.GetUnits().GetFormatOptions(spec).CanHaveSymbol(),
                                    ["SuppressSpaces"] = doc.GetUnits().GetFormatOptions(spec).SuppressSpaces,
                                    ["SuppressLeadingZeros"] = doc.GetUnits().GetFormatOptions(spec).SuppressLeadingZeros,
                                    ["SuppressTrailingZeros"] = doc.GetUnits().GetFormatOptions(spec).SuppressTrailingZeros,
                                    ["UsePlusPrefix"] = doc.GetUnits().GetFormatOptions(spec).UsePlusPrefix,

                                    // Accuracy has its own set of rules. Check the documentation on Autodesk.Revit.DB.FormatOptions.Accuracy to see valid ranges.
                                    ["SetAccuracy"] = new JObject
                                    {
                                        ["Value"] = doc.GetUnits().GetFormatOptions(spec).Accuracy,
                                        ["IsValidAccuracy"] = FormatOptions.IsValidAccuracy(unit, doc.GetUnits().GetFormatOptions(spec).Accuracy),
                                    },
                                    ["ValidSymbols"] = validSymbols(unit),
                                }

                            },
                            ["SetSymbol"] = new JObject
                            {
                                ["SymbolTypeId"] = symbol.TypeId,
                                ["SymbolLabel"] = symbolLabel(symbol),
                                ["IsValidSymbol"] = FormatOptions.IsValidSymbol(unit, symbol),
                            }
                        }
                    };
                    specData.Add(result);
                }
                catch (Exception ex)
                {
                    // Do any error handling here:
                    Debug.Print(ex.Message);
                    continue;
                }
            }

            // Wrap the data inside a JProperty named after the disciplineLabel.
            JProperty disciplineData = new JProperty(disciplineLabel, new JObject
            {
                // Attach the TypeId schema identifier.
                ["DisciplineTypeId"] = discipline.TypeId,
                // Add the JArray with all the spec data.
                ["Specs"] = specData
            });
            return disciplineData;
        }
    }
}