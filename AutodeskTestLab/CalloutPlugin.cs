using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System.Collections.Generic;
using System;
using System.Runtime.Remoting.Messaging;
using System.Linq;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.DatabaseServices.Filters;
using System.Data.SQLite;
using System.IO;

namespace AutodeskTestLab
{
    public class CalloutPlugin
    {
        // Основной метод плагина
        [CommandMethod("AutoLeaders")]
        public void Callout()
        {
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Editor acEd = acDoc.Editor;

            // Филтер выбора объектов
            PromptSelectionOptions optionSelect = new PromptSelectionOptions();
            SelectionFilter filterSelect = new SelectionFilter(
                new TypedValue[] { new TypedValue((int)DxfCode.Start, "INSERT") }
            );
            PromptSelectionResult selectionResult = acEd.GetSelection(optionSelect, filterSelect);

            // Проверка ввода пользователя
            switch (selectionResult.Status)
            {
                case PromptStatus.Cancel:
                    acEd.WriteMessage("Отмена выбора блока");
                    break;
                case PromptStatus.None:
                    acEd.WriteMessage("Не выбран блок");
                    break;
                case PromptStatus.Error:
                    acEd.WriteMessage("Ошибка!");
                    break;
                case PromptStatus.OK:
                    createMLeader(selectionResult);
                    break;
                default:
                    break;
            }

        }

        // Метод создания выносок
        private void createMLeader(PromptSelectionResult selectionResult)
        {
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            // Создание list для Id выбранных объектов
            List<ObjectId> objectIds = new List<ObjectId>();

            double maxPointBlockHeight = 0.0;

            // Начало транзакции для заполнение List objectIds
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;

                BlockTableRecord acBlkTblRec;
                acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;


                foreach (ObjectId acObjId in selectionResult.Value.GetObjectIds())
                {
                    DBObject acDbObj = acTrans.GetObject(acObjId, OpenMode.ForRead) as DBObject;
                    if (!(acDbObj == null || acDbObj is MLeader))
                    {
                        if (acDbObj is Entity acEntity)
                        {
                            Extents3d extents = acEntity.GeometricExtents;
                            maxPointBlockHeight = Math.Max(maxPointBlockHeight, extents.MaxPoint.Y);
                            objectIds.Add(acObjId);
                        }
                    }
                }
                // Сортировка по Handle.Value всех полученных объектов с чертежа
                objectIds.Sort(delegate(ObjectId x, ObjectId y)
                {
                    if (x.Handle.Value > y.Handle.Value) return 1;
                    else if (x.Handle.Value == y.Handle.Value) return 0; 
                    else return -1;    
                });

                
                acTrans.Commit();
            }

            // Транзакция для отрисовки выносок объектов
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;

                BlockTableRecord acBlkTblRec;
                acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                foreach (ObjectId acObjId in objectIds)
                {
                    DBObject acDbObj = acTrans.GetObject(acObjId, OpenMode.ForRead);

                    if (acDbObj != null)
                    {
                        if (acDbObj is Entity acEntity)
                        {

                            Point3d objectCenterPoint = acEntity.GeometricExtents.MinPoint + ((acEntity.GeometricExtents.MaxPoint - acEntity.GeometricExtents.MinPoint) / 2.0);

                            string contentNumberBlc = Convert.ToString(objectIds.IndexOf(acObjId) + 1);

                            if (objectIds.Any(o => (acTrans.GetObject(o, OpenMode.ForRead)).XData == acDbObj.XData))
                            {
                                ObjectId objectIdRep = objectIds.Find(o => (acTrans.GetObject(o, OpenMode.ForRead)).XData == acDbObj.XData);
                                contentNumberBlc = Convert.ToString(objectIds.IndexOf(objectIdRep) + 1);
                            }

                            // Создание MText для выноски
                            MText mText = new MText();
                            mText.Contents = contentNumberBlc;
                            mText.Location = new Point3d(objectCenterPoint.X + 30.0, maxPointBlockHeight + 30.0, 0);
                            mText.TextHeight = 15.0;

                            // Создание выноски с MText
                            MLeader mLeader = new MLeader();
                            mLeader.SetDatabaseDefaults();
                            mLeader.ContentType = ContentType.MTextContent;
                            mLeader.MText = mText;
                            mLeader.AddLeaderLine(objectCenterPoint);

                            acBlkTblRec.AppendEntity(mLeader);
                            acTrans.AddNewlyCreatedDBObject(mLeader, true);
                        }
                    }
                }

                createTable(objectIds);

                acTrans.Commit();
            }
        }
        // Метод создания таблицы
        private void createTable(List<ObjectId> objectIds)
        {
            Document acDoc = Application.DocumentManager.MdiActiveDocument;

            Database acCurDb = acDoc.Database;
            Editor ed = acDoc.Editor;

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {

                BlockTable acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;

                BlockTableRecord acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                double maxPlaneX = objectIds.Max(o => (acTrans.GetObject(o, OpenMode.ForRead) as Entity).GeometricExtents.MaxPoint.X);
                double minPlaneY = objectIds.Min(o => (acTrans.GetObject(o, OpenMode.ForRead) as Entity).GeometricExtents.MaxPoint.Y);

                Table acTable = new Table();
                acTable.SetDatabaseDefaults();
                acTable.TableStyle = acCurDb.Tablestyle;
                acTable.Position = new Point3d(maxPlaneX + 1000.0, minPlaneY, 0);


                int rowsNum = GetDataDb(objectIds).GroupBy(b => b.Id).ToList().Count();
                int columnNum = 4;
                double rowHeight = 50.0;
                double columnWidth = 40.0;

                acTable.InsertRows(0, rowHeight, rowsNum+1);
                acTable.InsertColumns(0, columnWidth, columnNum);

                acTable.SetRowHeight(rowHeight);
                acTable.SetColumnWidth(columnWidth);

                int[] customColumnWidths = new int[] { 200, 500, 200, 200 };
                string[] headlinesTable = new string[] { "№", "Наименование", "Количество", "Масса ед, кг" };

                // Обработка и подготовка полученных данных с базы данных SQLite
                var contentBlock = GetDataDb(objectIds)
                    .GroupBy(o => o.Id)
                    .Select(g => new
                    {
                        Id = g.Key,
                        Count = g.Count(),
                        Content = g.Select(p => new
                        {
                            Name = p.FullNameTemplate.ToString()
                            .Replace("<$D$>", p.Diameter.ToString())
                            .Replace("<$D1$>", p.Diameter.ToString()),
                            p.Weight
                        })
                    }).ToList();

                List<string[]> contentTable = new List<string[]>();
                foreach(var item in contentBlock)
                {
                    contentTable.Add(new string[]{
                        item.Content.Select(c => c.Name).ToArray()[0],
                        item.Count.ToString(),
                        item.Content.Select(c=> c.Weight).ToArray()[0].ToString()
                    });
                }
                contentTable.ToArray();

                // Заполнение таблицы данными
                for (int i = 0; i <= rowsNum; i++)
                {
                    for(int j = 0; j < columnNum; j++)
                    {
                        acTable.Cells[i, j].Alignment = CellAlignment.MiddleCenter;
                        acTable.Columns[j].Width = customColumnWidths[j];
                        acTable.Cells[i, j].TextHeight = 15;
                        if(i == 0)
                        {
                            acTable.Cells[i, j].TextString = headlinesTable[j].ToString();
                            continue;
                        }
                        if (j == 0)
                        {
                            acTable.Cells[i, 0].TextString = i.ToString();
                            continue;
                        }else if (i >= 1 && j != 0)
                        {
                            acTable.Cells[i, j].TextString = contentTable[i-1][j-1];
                        }
                    } 
                }


                acTable.GenerateLayout();
                acBlkTblRec.AppendEntity(acTable);
                acTrans.AddNewlyCreatedDBObject(acTable, true);
                acTrans.Commit();

            }
        }

        // Получение данных с базы данных SQLite
        private List<BlockEnt> GetDataDb(List<ObjectId> objectIds)
        {
            List<BlockEnt> blockEnts = new List<BlockEnt>();

            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                string appPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                string databasePath = "PartsDataBase.sqlite";
                string connectionString = $"Data Source={appPath}\\{databasePath};Version=3;";

                using (SQLiteConnection connection = new SQLiteConnection(connectionString))
                {
                    connection.Open();

                    foreach (ObjectId objId in objectIds)
                    {
                        BlockReference acBlocRef = acTrans.GetObject(objId, OpenMode.ForRead) as BlockReference;
                        string blockXDataId = acBlocRef.XData.AsArray()[2].Value.ToString();
                        string selectQuery = "SELECT * FROM 'Parts' WHERE ID = " + '"' + blockXDataId + '"' + ";";
                        using (SQLiteCommand command = new SQLiteCommand(selectQuery, connection))
                        {
                            using (SQLiteDataReader reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    BlockEnt block = new BlockEnt(
                                        reader.GetValue(0).ToString(),
                                        Convert.ToDouble(reader.GetValue(1)),
                                        Convert.ToInt32(reader.GetValue(2)),
                                        reader.GetValue(3).ToString());

                                    blockEnts.Add(block);
                                }
                            }
                        }
                    }
                }
            }
            return blockEnts;
        }
    }

}

