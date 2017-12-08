using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Bentley.IModel.Core;
using Bentley.IModel.Core.MetaData;
using NImodel.Utils;

namespace NImodel
{
    class IModelSchemax
    {
        public void schemainfo(string imodelpath)
        {
            IModel imodel = IModel.Open(imodelpath);
            string strInfo = "";
            foreach (Schema sch in imodel.Schemas)
            {
                strInfo += ("schemaName = " + sch.Name + "\n");

                foreach (Class cl in sch.Classes)
                {
                    strInfo += ("\tclass = " + cl.Name + "\n");

                    foreach (Property pr in cl.Properties)
                    {
                        strInfo += ("\t\tProperty = " + pr.CLRType.Name + " " + pr.Name + "\n");
                    }

                    foreach (Relationship rel in cl.Relationships)
                    {
                        strInfo += ("\t\trelationship = " + rel.Source.Name + "->" + rel.Target.Name + "\n");
                    }
                }

                foreach (string strRela in sch.Relationships)
                {
                    strInfo += ("\tRelationship = " + strRela + "\n");
                }

                strInfo += ("\n\n");
            }
            //Console.WriteLine(strInfo);
            FileIOUtil.writeToFile(strInfo);
        }
    }
}
