using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using AcApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;

namespace AutoTableFromDataLink;

public sealed class Commands
{
    private const double CellWidth = 2025.0;
    private const double CellHeight = 486.0;
    private const double TextHeight = 200.0;

    [CommandMethod("AUTOTABLEDATALINK", CommandFlags.Modal)]
    public void AutoTableDataLink()
    {
        Run();
    }

    [CommandMethod("AUTOTABLEDL", CommandFlags.Modal)]
    public void AutoTableDataLinkShort()
    {
        Run();
    }

    private static void Run()
    {
        Document? doc = AcApp.DocumentManager.MdiActiveDocument;
        if (doc == null)
        {
            return;
        }

        Database db = doc.Database;
        Editor ed = doc.Editor;

        try
        {
            using (doc.LockDocument())
            {
                DataLinkChoice? choice = PromptForDataLink(ed, db);
                if (choice == null)
                {
                    return;
                }

                PromptPointResult pointResult = ed.GetPoint(
                    new PromptPointOptions("\n選擇表格插入位置: "));
                if (pointResult.Status != PromptStatus.OK)
                {
                    return;
                }

                ObjectId tableId = CreateLinkedTable(db, choice, pointResult.Value);
                ed.Regen();

                ed.WriteMessage(
                    "\n已建立表格：資料連結「{0}」，寬度 {1}，高度 {2}，文字高度 {3}。",
                    choice.Name,
                    CellWidth,
                    CellHeight,
                    TextHeight);
                ed.WriteMessage("\nTable ObjectId: {0}", tableId);
            }
        }
        catch (Autodesk.AutoCAD.Runtime.Exception ex)
        {
            ed.WriteMessage("\nAutoCAD 錯誤: {0}", ex.Message);
        }
        catch (Exception ex)
        {
            ed.WriteMessage("\n外掛錯誤: {0}", ex.Message);
        }
    }

    private static DataLinkChoice? PromptForDataLink(Editor ed, Database db)
    {
        List<DataLinkChoice> links = GetDataLinks(db);
        if (links.Count == 0)
        {
            ed.WriteMessage("\n目前圖面沒有資料連結。請先用 DATALINK 或 AUTODATALINKSHEETS 建立資料連結。");
            return null;
        }

        ed.WriteMessage("\n可用資料連結:");
        for (int i = 0; i < links.Count; i++)
        {
            ed.WriteMessage("\n  {0}. {1}", i + 1, links[i].Name);
        }

        PromptIntegerOptions options = new("\n選資料連結編號: ")
        {
            AllowNone = false,
            LowerLimit = 1,
            UpperLimit = links.Count,
        };

        PromptIntegerResult result = ed.GetInteger(options);
        return result.Status == PromptStatus.OK ? links[result.Value - 1] : null;
    }

    private static List<DataLinkChoice> GetDataLinks(Database db)
    {
        List<DataLinkChoice> links = new();

        ObjectId dataLinkDictionaryId = db.DataLinkDictionaryId;
        if (dataLinkDictionaryId.IsNull)
        {
            return links;
        }

        using Transaction tr = db.TransactionManager.StartTransaction();
        DBDictionary dictionary = (DBDictionary)tr.GetObject(dataLinkDictionaryId, OpenMode.ForRead);

        foreach (DBDictionaryEntry entry in dictionary)
        {
            try
            {
                DataLink? dataLink = tr.GetObject(entry.Value, OpenMode.ForRead) as DataLink;
                if (dataLink == null || dataLink.IsErased)
                {
                    continue;
                }

                string name = string.IsNullOrWhiteSpace(dataLink.Name) ? entry.Key : dataLink.Name;
                links.Add(new DataLinkChoice(name, entry.Value));
            }
            catch (Autodesk.AutoCAD.Runtime.Exception)
            {
                // Ignore stale dictionary entries and keep listing usable links.
            }
        }

        tr.Commit();
        links.Sort((left, right) => string.Compare(left.Name, right.Name, StringComparison.CurrentCultureIgnoreCase));
        return links;
    }

    private static ObjectId CreateLinkedTable(Database db, DataLinkChoice choice, Point3d insertionPoint)
    {
        using Transaction tr = db.TransactionManager.StartTransaction();

        DataLink dataLink = (DataLink)tr.GetObject(choice.Id, OpenMode.ForWrite);
        dataLink.Update(UpdateDirection.SourceToData, UpdateOption.None);

        BlockTableRecord currentSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

        Table table = new()
        {
            TableStyle = db.Tablestyle,
            Position = insertionPoint,
        };

        table.SetSize(1, 1);
        ApplyTableFormat(table);

        table.SetDataLink(0, 0, choice.Id, true);
        ObjectId tableId = currentSpace.AppendEntity(table);
        tr.AddNewlyCreatedDBObject(table, true);

        table.UpdateDataLink(UpdateDirection.SourceToData, UpdateOption.None);
        ApplyTableFormat(table);
        table.GenerateLayout();

        tr.Commit();
        return tableId;
    }

    private static void ApplyTableFormat(Table table)
    {
        table.SetColumnWidth(CellWidth);
        table.SetRowHeight(CellHeight);

        for (int row = 0; row < table.NumRows; row++)
        {
            for (int column = 0; column < table.NumColumns; column++)
            {
                table.SetTextHeight(row, column, TextHeight);
            }
        }
    }

    private sealed class DataLinkChoice
    {
        public DataLinkChoice(string name, ObjectId id)
        {
            Name = name;
            Id = id;
        }

        public string Name { get; }

        public ObjectId Id { get; }
    }
}
