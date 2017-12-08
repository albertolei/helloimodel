using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Bentley.IModel.Core;
using Bentley.IModel.Core.MetaData;
using Bentley.IModel.Core.BusinessData;

namespace NImodel.Utils
{
    class ModifyUtil
    {

        public void addDynamicToIModel(string imodelPath, string schemaPath, string ecclassType)
        {
            IModel imodel = IModel.OpenForAugmentation(imodelPath, imodelPath.Replace(".i", "new.i"));
            Console.WriteLine(imodel == null);
            Schema schema = null;
            try
            {
                schema = imodel.ImportSchema(schemaPath, "");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }

            if (imodel == null)
            {
                Console.WriteLine("Could not create i-model: imodel is null");
            }
            else if(schema == null)
            {
                Console.WriteLine("Could not create i-model: schema is null");
            }
            else
            {
                dynamic d0 = new Dynamics(imodel, schema[ecclassType]);
                d0.CatalogTypeName = "对象类型：test";
                d0.CatalogInstanceName = "对象实例:test实例";
                Console.WriteLine(d0 == null);
                
                IModelElement element0 = imodel.CreateElement();
                Console.WriteLine(element0 == null);

                try
                {
                    element0.Add(d0);
                    imodel.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    Console.WriteLine(ex.StackTrace);
                }
            }
        }

        public void addDynamicsToIModelElement(string imodelPath, string schemaPath, string ecclassType, string elementId)
        {
            IModel imodel = IModel.OpenForAugmentation(imodelPath, imodelPath.Replace(".i", "new.i"));
            IModelElement element = imodel.GetElementByID(long.Parse(elementId));
            try
            {
                Schema schema = imodel.ImportSchema(schemaPath, "");
                dynamic d0 = new Dynamics(imodel, schema[ecclassType]);
                d0.CatalogTypeName = "对象类型：test";
                d0.CatalogInstanceName = "对象实例:test实例";
                element.Add(d0);
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine(e.StackTrace);
            }
            finally
            {
                imodel.Close();
            }
            
        }
    }

}
