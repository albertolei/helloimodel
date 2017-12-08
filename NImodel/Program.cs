using System;
using System.IO;
using System.Text;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Drawing;
using Bentley.IModel.Core;
using Bentley.IModel.Core.BusinessData;
using Bentley.IModel.Core.MetaData;
using Bentley.GeometryNET.Common;
using Bentley.DgnPlatformNET.Elements;
using Bentley.DgnPlatformNET;
using NImodel.Utils;
using NImodel.Dao;

namespace NImodel
{
    class Program
    {
        private string dataSource = @"D:\Access\schema.accdb";
        private static string imodelFileName = @"D:\idgn\zhouwang.i.dgn";

        static void Main(string[] args)
        {
            Program p = new Program();
            p.test(imodelFileName);
            Console.Read();
        }

        public void test(string filename)
        {
            IModel imodel = IModel.Open(filename);
            List<Schema> schemas = imodel.Schemas;
            foreach (Schema schema in schemas)
            {
                schema.SaveAs(@"d:\imodel_info\" + schema.Name + ".xml");
            }
            IEnumerable<IModelElement> elements = imodel.Elements;
            foreach (IModelElement element in elements)
            {
                FileIOUtil.writeToFile("元素ID:" + element.Element.ElementId);
                FileIOUtil.writeToFile("元素类型[element.Element.ElementType]:" + element.Element.ElementType.ToString());
                FileIOUtil.writeToFile("元素层次:" + element.LevelName);
                FileIOUtil.writeToFile("元素类型[element.TypeName]:" + element.TypeName);
                FileIOUtil.writeToFile("元素颜色：" + element.Color.ToString());
                FileIOUtil.writeToFile("元素描述:" + element.Description);
                FileIOUtil.writeToFile("元素线宽:" + element.LineWeight);

                FileIOUtil.writeToFile("elementclass_num:" + element.Classes.Count);
                foreach (Class c in element.Classes)
                {
                    FileIOUtil.writeToFile("元素分类:" + c.Name);
                }
                FileIOUtil.writeToFile("\n");
                foreach (Dynamics o in element.Objects)
                {
                    foreach (Property p in o.Class.Properties)
                    {
                        if (!o.ECInstance[p.Name].IsNull)
                        {
                            FileIOUtil.writeToFile("Objects属性名:" + p.Name + ",Objects属性值:" + o.ECInstance[p.Name].NativeValue.ToString());
                        }
                    }
                }
                FileIOUtil.writeToFile("\n");

            }
            imodel.Close();
        }

        public void test1()
        {
            //Program p = new Program();
            //p.test();

            //p.imodelparser(@"d:\testvba.i.dgn");
            //Console.Read();


            //List<string> names = new List<string>();
            //names.Add("aname");
            //names.Add("apassword");
            //List<string> values = new List<string>();
            //values.Add("shan");
            //values.Add("654321");
            //ModelDao modelDao = new ModelDao();
            //modelDao.write("table1", names, values);

            string filename = @"d:\3132\我的文档\database1.accdb";
            string tablename = "table4";
            ModelDao modelDao = new ModelDao();
            bool isExist = modelDao.isTableExist(filename, tablename);
            if (isExist)
            {
                Console.WriteLine("table exists");
            }
            else
            {
                bool result = modelDao.createTable(filename, tablename);
                Console.WriteLine(result);
            }


            //bool isFieldExist = modelDao.isFieldExist(@"d:\3132\我的文档\database1.accdb", "table1","aname");
            //Console.WriteLine(isFieldExist);
            //if (modelDao.isFieldExist(@"d:\3132\我的文档\database1.accdb", "table1", "tname") == 0)
            //{
            //    bool result = modelDao.createField(@"d:\3132\我的文档\database1.accdb", "table1", "tname", "varchar");
            //    if (result == false)
            //    {
            //        Console.WriteLine("fail");
            //    }
            //    else
            //    {
            //        Console.WriteLine("success");
            //    }
            //}
            //else 
            //{
            //    Console.WriteLine("表不已存在或表中已经存在该字段");
            //} 
        }

        public void testChildrenElement()
        {
            IModel imodel = IModel.Open(imodelFileName);
            IModelElement imodelElement = imodel.GetElementByID(68556);
            Console.WriteLine(imodelElement.TypeName);
            IEnumerable<IModelElement> children = imodelElement.Children;
            Console.WriteLine(children.Count());

            ChildElementCollection cec = imodelElement.Element.GetChildren();
            Console.WriteLine(cec.Count());

            //Console.WriteLine(imodelElement == null);
            //IEnumerable<IModelElement> elements = imodel.Elements;
            //foreach (IModelElement element in elements)
            //{
                //ModelElementsCollection mec = element.Element.DgnModel.GetElements();

                

                //getChild(element.Element);
                //Console.WriteLine(element.Element.ElementId.ToString());
                //Console.WriteLine(element.Element.ElementId.ToString() == "68556");
                //if (element.Element.ElementId.ToString() == "68556")
                //{
                    //long id = 68571;
                    //Element ele68571 = element.Element.DgnModel.FindElementById(new ElementId(ref id));
                    //Console.WriteLine(ele68571 == null);
                    //Console.WriteLine(element.Element.GetChildren().Count() > 0);
                    //if (element.Children.Count() > 0)
                    //{
                    //    IEnumerable<IModelElement> children = element.Children;
                    //    foreach (IModelElement child in children)
                    //    {
                    //        if (child.Element.ElementId.ToString() == "68572")
                    //        {
                    //            IModelElement target = child;
                    //            foreach (Dynamics obj in target.Objects)
                    //            {
                    //                foreach (Property prop in obj.Class.Properties)
                    //                {
                    //                    Console.WriteLine(prop.Name + ":" + obj.ECInstance[prop.Name]);
                    //                }
                    //            }
                    //        }
                    //    }
                    //}
                //}
            //}
        }

        private void getChild(Element element)
        {
            FileIOUtil.writeToFile(element.ElementId.ToString());
            if (element.DgnModel.GetElements().Count() > 0)
            {
                foreach (Element child in element.DgnModel.GetElements())
                {
                    getChild(child);
                }
            }
           
        }
        
        //解析imodel自身的属性
        public void imodelparser(string filename)
        {
            string tablename = "imodel";
            List<string> keys = new List<string>();
            List<string> values = new List<string>();
            IModel iModel = IModel.Open(filename);


            keys.Add("ActiveModel");
            values.Add(iModel.ActiveModel);

            keys.Add("IsCapturingRelations");
            values.Add(iModel.IsCapturingRelations+"");

            keys.Add("IsSecure");
            values.Add(iModel.IsSecure+"");

            keys.Add("Name");
            values.Add(iModel.Name);

            ModelDao modelDao = new ModelDao();
            if (!modelDao.isTableExist(dataSource, tablename))
            {
                modelDao.createTable(dataSource, tablename);
            }

            foreach (string key in keys)
            {
                if (modelDao.isFieldExist(dataSource, tablename, key) == 0)
                {
                    modelDao.createField(dataSource, tablename, key, "varchar");
                }
            }

            modelDao.write(dataSource, tablename,keys, values);
            
        }
        
        //imodel.classes的属性
        public void imodelClassParser(string filename)
        {
            IModel iModel = IModel.Open(filename);

            //IEnumerable<Dynamics> objects =  iModel.GetObjects();
            string modelName = iModel.Name;
            List<Class> cs = iModel.Classes;
            
            int i = 0;
            foreach (Class c in cs)
            {
                FileIOUtil.writeToFile("第" + i++ + "个class:");
                FileIOUtil.writeToFile("属性名：c.Name，" + "属性值：" + c.Name);
                FileIOUtil.writeToFile("属性名：c.ECClass.Name，" + "属性值：" + c.ECClass.Name);
                FileIOUtil.writeToFile("属性名：c.ECClass.CustomStructSerializerName，" + "属性值：" + c.ECClass.CustomStructSerializerName);
                FileIOUtil.writeToFile("属性名：c.ECClass.Description，" + "属性值：" + c.ECClass.Description);
                FileIOUtil.writeToFile("属性名：c.ECClass.DisplayLabel，" + "属性值：" + c.ECClass.DisplayLabel);
                FileIOUtil.writeToFile("属性名：c.ECClass.IsCustomAttribute，" + "属性值：" + c.ECClass.IsCustomAttribute);
                FileIOUtil.writeToFile("属性名：c.ECClass.IsDisplayLabelDefined，" + "属性值：" + c.ECClass.IsDisplayLabelDefined);
                FileIOUtil.writeToFile("属性名：c.ECClass.IsDomainClass，" + "属性值：" + c.ECClass.IsDomainClass);
                FileIOUtil.writeToFile("属性名：c.ECClass.IsStruct，" + "属性值：" + c.ECClass.IsStruct);
                FileIOUtil.writeToFile("属性名：c.ECClass.RequiresCustomStructSerializer，" + "属性值：" + c.ECClass.RequiresCustomStructSerializer);
                FileIOUtil.writeToFile("属性名：c.ECClass.Schema.Description，" + "属性值：" + c.ECClass.Schema.Description);
                FileIOUtil.writeToFile("\n");

                
            }
        }

        //将imodel.classes的属性输出到access数据库
        public void imodelClassParserToDB(string filename)
        {
            string tablename = "class";
            List<string> keys = new List<string>();
            List<List<string>> values = new List<List<string>>();
            IModel iModel = IModel.Open(filename);

            string modelName = iModel.Name;
            List<Class> cs = iModel.Classes;

            //keys.Add("Name");
            foreach(Class c in cs)
            {
                //List<string> value = new List<string>();
                //value.Add(c.Name);
                //values.Add(value);
                Console.WriteLine("c.Name:" + c.Name);
                Console.WriteLine("c.ECClass.Name:" + c.ECClass.Name);
                Console.WriteLine("CustomStructSerializerName:" + c.ECClass.CustomStructSerializerName);
                Console.WriteLine("Description:" + c.ECClass.Description);
                Console.WriteLine("DisplayLabel:" + c.ECClass.DisplayLabel);
                Console.WriteLine("IsCustomAttribute:" + c.ECClass.IsCustomAttribute);
                Console.WriteLine("IsDisplayLabelDefined:" + c.ECClass.IsDisplayLabelDefined);
                Console.WriteLine("IsDomainClass:" + c.ECClass.IsDomainClass);
                Console.WriteLine("IsStruct:" + c.ECClass.IsStruct);
                Console.WriteLine("RequiresCustomStructSerializer:" + c.ECClass.RequiresCustomStructSerializer);
                Console.WriteLine("Description:" + c.ECClass.Schema.Description);
                Console.WriteLine();
            }

            ModelDao modelDao = new ModelDao();
            if (!modelDao.isTableExist(dataSource, tablename))
            {
                modelDao.createTable(dataSource, tablename);
            }

            foreach (string key in keys)
            {
                if (modelDao.isFieldExist(dataSource, tablename, key) == 0)
                {
                    modelDao.createField(dataSource, tablename, key, "varchar");
                }
            }

            modelDao.insert(dataSource, tablename, keys, values);
        }

        //得到imodel的element拥有的dynamics
        public void imodelElmentsDynamicsParser(string filename)
        {
            IModel imodel = IModel.Open(filename);
            IEnumerable<IModelElement> elements = imodel.Elements;
            foreach (IModelElement element in elements)
            {
                Console.WriteLine(element.Element.ElementId);
                foreach(Dynamics d in element.Objects)
                {
                    Console.WriteLine(d.Class.Name);
                }
                Console.WriteLine();
            }
        }

        //得到imodel.Elements的geometry属性
        public void imodelElementsGeometry(string filename)
        {
            IModel imodel = IModel.Open(filename);
            IEnumerable<IModelElement> elements = imodel.Elements;
            foreach (IModelElement element in elements)
            {
                IGeometry geometry = element.Geometry;
                Console.WriteLine(geometry.GetType().ToString());
                
            }
        }


        //得到imodel.Elements的element属性
        public void imodelElementsParser(string filename)
        {
            IModel imodel = IModel.Open(imodelFileName);
            IEnumerable<IModelElement> elements = imodel.Elements;
            foreach (IModelElement element in elements)
            {
                /*关于imodel的element自身的属性 
                FileIOUtil.writeToFile("元素ID: " + element.Element.ElementId);
                FileIOUtil.writeToFile("element颜色: " + element.Color);

                if (element.TypeName.Equals("Cell"))
                {
                    FileIOUtil.writeToFile("element填充颜色: Cell元素没有填充颜色.");
                    FileIOUtil.writeToFile("element文字: Cell元素没有填充颜色.");
                }
                else
                {
                    if (element.TypeName.Equals("Line"))
                    {
                        FileIOUtil.writeToFile("element填充颜色: Line元素没有填充颜色");
                        FileIOUtil.writeToFile("element文字: Line元素没有文字");
                    }
                    else
                    {
                        FileIOUtil.writeToFile("element填充颜色: " + element.FillColor);
                        FileIOUtil.writeToFile("element文字: " + element.TextString);
                    }
                }

                FileIOUtil.writeToFile("element.Element类型[element.Element.ElementType]: " + element.Element.ElementType.ToString());
                FileIOUtil.writeToFile("element层次: " + element.LevelName);
                FileIOUtil.writeToFile("element类型[element.TypeName]: " + element.TypeName);
                FileIOUtil.writeToFile("element描述: " + element.Description);
                FileIOUtil.writeToFile("element线宽: " + element.LineWeight);
                FileIOUtil.writeToFile("element透明度: " + element.Transparency);

                FileIOUtil.writeToFile("element.Objects.Class.Properties属性: ");
                */
                foreach (Dynamics o in element.Objects)
                {
                    //返回当前dynamics对象o的类Bentley.IModel.Core.BusinessData.Dynamics
                    //FileIOUtil.writeToFile("element.objects.gettype().tostring(): " + o.GetType().ToString());
                    // 得到Dynamics中包含的属性，以及属性的属性
                    foreach (Property p in o.Class.Properties)
                    {
                        if (!o.ECInstance[p.Name].IsNull)
                        {
                            FileIOUtil.writeToFile("属性名: " + p.Name + ", 属性值: " + o.ECInstance[p.Name].NativeValue.ToString());
                            

                        }
                        FileIOUtil.writeToFile("p.Name: " + p.Name);
                        FileIOUtil.writeToFile("p.Catelog: " + p.Category);
                        FileIOUtil.writeToFile("p.IsArray: " + p.IsArray);
                        FileIOUtil.writeToFile("p.IsStruct: " + p.IsStruct);
                        FileIOUtil.writeToFile("p.CLRType: " + p.CLRType);
                        FileIOUtil.writeToFile("----\n");
                    }
                    

                    /*Dynamics的IECinstance属性
                    IECInstance ecInstance = o.ECInstance;
                    FileIOUtil.writeToFile("IECInstanceID: " + ecInstance.InstanceId);
                    FileIOUtil.writeToFile("IECInstanceIsReadOnly: " + ecInstance.IsReadOnly);
                    FileIOUtil.writeToFile("IECInstanceContainsValues: " + ecInstance.ContainsValues);
                    FileIOUtil.writeToFile("-----\n");
                    */


                    /*返回当前dynamics对象o的关联对象，但没有任何输出。
                    Dictionary<string, RelatedObjects> dic = o.RelatedObjects;
                    foreach(string key in dic.Keys)
                    {
                        Console.WriteLine(key);
                        RelatedObjects ros = dic[key];
                        Console.WriteLine(ros.Relationship.Name);
                        Console.WriteLine(ros.Relationship.Source);
                        Console.WriteLine(ros.Relationship.Target);
                        Console.WriteLine();
                    }
                    */
                }
                FileIOUtil.writeToFile("-----\n");
            }
            imodel.Close();
        }
        
        //将解析到的imodel.Elements属性写入一个Dictionary
        public Dictionary<string, Dictionary<string, string>> parseIModelElementsToDict(string filename)
        {
            IModel imodel = IModel.Open(imodelFileName);
            IEnumerable<IModelElement> elements = imodel.Elements;
            Dictionary<string, Dictionary<string, string>> imodelInfo = new Dictionary<string,Dictionary<string,string>>();
            foreach (IModelElement element in elements)
            {
                Dictionary<string, string> elementInfo = new Dictionary<string,string>();
                //Dynamics对应了一个ECClass
                foreach(Dynamics dynamics in element.Objects)
                {
                    foreach(Property pp in dynamics.Class.Properties)
                    {
                        if (!dynamics.ECInstance[pp.Name].IsNull)
                        {
                            elementInfo.Add(pp.Name, dynamics.ECInstance[pp.Name].NativeValue.ToString());
                        }
                    }
                }
                imodelInfo.Add(element.Element.ElementId.ToString(), elementInfo);
            }
            return imodelInfo;
        }

        //得到imodel.Elements的class属性
        public void imodelElementsClassParser(string filename)
        {
            IModel imodel = IModel.Open(imodelFileName);
            IEnumerable<IModelElement> elements = imodel.Elements;
            int k = 0;
            foreach (IModelElement element in elements)
            {
                FileIOUtil.writeToFile("第" + k++ + "个Element的class属性:");
                List<Class> cs = element.Classes;
                int i = 0;
                foreach (Class c in cs)
                {
                    foreach (Property p in c.Properties)
                    {
                        FileIOUtil.writeToFile("p.Name: " + p.Name);
                        FileIOUtil.writeToFile("p.Catelog: " + p.Category);
                        FileIOUtil.writeToFile("p.IsArray: " + p.IsArray);
                        FileIOUtil.writeToFile("p.IsStruct: " + p.IsStruct);
                        FileIOUtil.writeToFile("p.CLRType: " + p.CLRType);
                        FileIOUtil.writeToFile("---\n");
                    }
                    FileIOUtil.writeToFile("第" + i++ + "个class:");
                    //FileIOUtil.writeToFile("属性名：c.Name，" + "属性值：" + c.Name);
                    //FileIOUtil.writeToFile("属性名：c.ECClass.Name，" + "属性值：" + c.ECClass.Name);
                    //FileIOUtil.writeToFile("属性名：c.ECClass.CustomStructSerializerName，" + "属性值：" + c.ECClass.CustomStructSerializerName);
                    //FileIOUtil.writeToFile("属性名：c.ECClass.Description，" + "属性值：" + c.ECClass.Description);
                    //FileIOUtil.writeToFile("属性名：c.ECClass.DisplayLabel，" + "属性值：" + c.ECClass.DisplayLabel);
                    //FileIOUtil.writeToFile("属性名：c.ECClass.IsCustomAttribute，" + "属性值：" + c.ECClass.IsCustomAttribute);
                    //FileIOUtil.writeToFile("属性名：c.ECClass.IsDisplayLabelDefined，" + "属性值：" + c.ECClass.IsDisplayLabelDefined);
                    //FileIOUtil.writeToFile("属性名：c.ECClass.IsDomainClass，" + "属性值：" + c.ECClass.IsDomainClass);
                    //FileIOUtil.writeToFile("属性名：c.ECClass.IsStruct，" + "属性值：" + c.ECClass.IsStruct);
                    //FileIOUtil.writeToFile("属性名：c.ECClass.RequiresCustomStructSerializer，" + "属性值：" + c.ECClass.RequiresCustomStructSerializer);
                    //FileIOUtil.writeToFile("属性名：c.ECClass.Schema.Description，" + "属性值：" + c.ECClass.Schema.Description);
                    FileIOUtil.writeToFile("----\n");
                }
                FileIOUtil.writeToFile("-----\n");
            }
        }

        //得到imodel.schemas，导出成xml文件
        public void imodelSchemasParser(string filename)
        {
            IModel imodel = IModel.Open(filename);
            //IModelView view = imodel.Views;
            //try 
            //{
            //    Console.WriteLine(view == null);
            //    Console.WriteLine(view.Name);
            //}
            //catch (Exception e)
            //{
            //    Console.WriteLine(e.Message);
            //    Console.WriteLine(e.StackTrace);
            //}
            List<Schema> schemas = imodel.Schemas;
            foreach (Schema schema in schemas)
            {
                schema.SaveAs(@"d:\imodel_info\" + schema.Name + ".xml");
                IEnumerable<Class> classes = schema.Classes;
                foreach (Class c in classes)
                {
                    Console.WriteLine(c.ECClass.Name);

                }
                Console.WriteLine("----");
            }
            Console.WriteLine("-------");
        }

        //得到imodel.documents
        public void imodelDocumentsParser(string filename)
        {
            IModel imodel = IModel.Open(filename);
            List<EmbeddedDocument> docs = imodel.Documents;
            Console.WriteLine(docs.Count);
            foreach (EmbeddedDocument doc in docs)
            {
                if (doc.Retrieve())
                {
                    Console.WriteLine("success");
                }
                else
                {
                    Console.WriteLine("false");
                }
            }
        }

        //得到imodel的tag元素
        public void imodelTagElementParser(string filename)
        {
            IModel imodel = IModel.Open(filename);
            foreach (IModelElement element in imodel.Elements)
            {
                if (element.Element.ElementType == MSElementType.Tag)
                {
                    TagElement tagEle = (TagElement)element.Element;
                    string tagValue = tagEle.GetTagValue().ToString();
                    Console.WriteLine(tagValue);
                }
            }
        }
    }
}
