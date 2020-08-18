using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ML
{
    public static class ComLibClass
    {
        /// <summary>
        /// 获取当前.NET程序所在的目录
        /// </summary>
        /// <returns>返回当前.NET程序所在的目录</returns>
        public static string GetCurrentPath()
        {
            var moudle = System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0];
            return System.IO.Path.GetDirectoryName(moudle.FullyQualifiedName);
        }
        /// <summary>
        /// 从外部DWG文件中插入图块
        /// </summary>
        /// <param name="db">图形数据库</param>
        /// <param name="dwgFileName">外部DWG文件名全称</param>
        /// <param name="blkname">块名</param>
        public static void ImportBlocksFromDWG(this Database db, string dwgFileName, string blkname)
        {
            Database tempDB = new Database(false, true);//新建一个临时图形数据库<1>
            try
            {
                tempDB.ReadDwgFile(dwgFileName, System.IO.FileShare.Read, true, null);//读取外部DWG到临时数据库<2>
                ObjectIdCollection ids = new ObjectIdCollection();//存放块OjbectId的集合
                using (Transaction trans = tempDB.TransactionManager.StartTransaction())
                {
                    BlockTable bt = trans.GetObject(tempDB.BlockTableId, OpenMode.ForRead) as BlockTable;
                    if (bt.Has(blkname))
                    {
                        BlockTableRecord btr = trans.GetObject(bt[blkname], OpenMode.ForRead) as BlockTableRecord;
                        ids.Add(bt[blkname]);
                    }
                }
                tempDB.WblockCloneObjects(ids, db.BlockTableId, new IdMapping(), DuplicateRecordCloning.Replace, false);
            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                Application.ShowAlertDialog("\n错误：" + ex.Message);
            }
            tempDB.Dispose();
        }

        /// <summary>
        /// 用于在AutoCAD图形中插入块参照
        /// </summary>
        /// <param name="spaceId">块参照要加入的模型空间或图纸空间的Id</param>
        /// <param name="layer">块参照要加入的图层名</param>
        /// <param name="blockName">块参照所属的块名</param>
        /// <param name="position">块参照的插入点</param>
        /// <param name="scale">块参照的缩放比例</param>
        /// <param name="rotateAngle">块参照的旋转角度</param>
        /// <returns></returns>
        public static ObjectId InsertBlockReference(this ObjectId spaceId, string layer, string blockName, Point3d position, Scale3d scale, double rotateAngle)
        {
            ObjectId blockRefId;
            Database db = spaceId.Database;
            //打开块表
            BlockTable bt = (BlockTable)db.BlockTableId.GetObject(OpenMode.ForRead);
            if (!bt.Has(blockName)) return ObjectId.Null;
            //以写的方式打开空间（模型空间或图纸空间）
            BlockTableRecord space = (BlockTableRecord)spaceId.GetObject(OpenMode.ForWrite);
            //创建一个块参照并设置插入点
            BlockReference br = new BlockReference(position, bt[blockName]);
            //设置块参照的缩放比例
            br.ScaleFactors = scale;
            //设置块参照的层名
            br.Layer = layer;
            //设置块参照的旋转角度
            br.Rotation = rotateAngle;
            //在空间中加入创建的块参照
            blockRefId = space.AppendEntity(br);
            //通知事务处理加入创建的块参照
            db.TransactionManager.AddNewlyCreatedDBObject(br, true);
            space.DowngradeOpen();
            return blockRefId;
        }

        /// <summary>
        /// 用于在AutoCAD图形中插入带属性的块参照
        /// </summary>
        /// <param name="spaceId">块参照要加入的模型空间或图纸空间的Id</param>
        /// <param name="layer">块参照要加入的图层名</param>
        /// <param name="blockName">块参照所属的块名</param>
        /// <param name="position">块参照的插入点</param>
        /// <param name="scale">块参照的缩放比例</param>
        /// <param name="rotateAngle">块参照的旋转角度</param>
        /// <param name="attNameValues">属性的名称和取值</param>
        /// <returns></returns>
        public static ObjectId InsertBlockReference(this ObjectId spaceId, string layer, string blockName, Point3d position, Scale3d scale, double rotateAngle, Dictionary<string, string> attNameValues)
        {
            //获取数据库对象
            Database db = spaceId.Database;
            //以读的方式打开块表
            BlockTable bt = (BlockTable)db.BlockTableId.GetObject(OpenMode.ForRead);
            //如果没有blockName表示的块，则程序返回
            if (!bt.Has(blockName)) return ObjectId.Null;
            //以写的方式打开空间(模型空间或图纸空间)
            BlockTableRecord space = (BlockTableRecord)spaceId.GetObject(OpenMode.ForWrite);
            //获取块表记录的Id
            ObjectId btrId = bt[blockName];
            //打开块表记录
            BlockTableRecord btr = (BlockTableRecord)btrId.GetObject(OpenMode.ForRead);
            //创建一个块参照并设置插入点
            BlockReference br = new BlockReference(position, bt[blockName]);
            //设置块参照的缩放比例
            br.ScaleFactors = scale;
            //设置块参照的图层
            br.Layer = layer;
            //设置块参照的旋转角度
            br.Rotation = rotateAngle;
            //将块参照添加到模型空间
            space.AppendEntity(br);

            //判断块表记录是否包含属性定义
            if (btr.HasAttributeDefinitions)
            {
                //遍历属性定义
                foreach (ObjectId id in btr)
                {
                    //检查是否是属性定义
                    AttributeDefinition attDef = id.GetObject(OpenMode.ForRead) as AttributeDefinition;
                    if (attDef != null)
                    {
                        //创建一个新的属性对象
                        AttributeReference attribute = new AttributeReference();
                        //从属性定义获得属性对象的对象特征
                        attribute.SetAttributeFromBlock(attDef, br.BlockTransform);
                        //设置属性对象的其它特征
                        attribute.Position = attDef.Position.TransformBy(br.BlockTransform);
                        attribute.Rotation = attDef.Rotation;
                        attribute.AdjustAlignment(db);
                        //判断是否包含指定的属性名称
                        if (attNameValues.ContainsKey(attDef.Tag.ToUpper()))
                        {
                            //设置属性值
                            attribute.TextString = attNameValues[attDef.Tag.ToUpper()].ToString();
                        }
                        //向块参照添加属性对象
                        br.AttributeCollection.AppendAttribute(attribute);
                        db.TransactionManager.AddNewlyCreatedDBObject(attribute, true);
                    }
                }
            }
            db.TransactionManager.AddNewlyCreatedDBObject(br, true);
            //返回添加的块参照的Id
            return br.ObjectId;
        }

    }
}
