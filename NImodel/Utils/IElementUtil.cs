using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Bentley.IModel.Core;
using Bentley.IModel.Core.BusinessData;
using System.IO;

namespace NImodel.Utils
{
    class IElementUtil
    {
        private static readonly string catalogPath = "../../CatalogTypes.txt";
        private static HashSet<string> catalogtypes = new HashSet<string>();
        private static readonly HashSet<string> tartgetProps = new HashSet<string> { "SECTNAME", "PTS_0", "PTS_1", "RORATION", "LENGTH", "OV", "CatalogTypeName", "CatalogInstanceName", "RangeLow", "RangeHigh", "Origin", "Slab_Common__x002f____x0040__Thickness", "SlabQuantities__x002f____x0040__Bottom", "SlabQuantities__x002f____x0040__Top"};

        public HashSet<IModelElement> getStructElements(string imodelpath)
        {
            initCatalogTypes();
            HashSet<IModelElement> ret = new HashSet<IModelElement>();
            IModel imodel = IModel.Open(imodelpath);
            foreach (IModelElement element in imodel.Elements)
            {
                foreach (Dynamics o in element.Objects)
                {
                    if (o.GetDynamicMemberNames().Contains("CatalogTypeName") && isStruct(o.ECInstance["CatalogTypeName"].NativeValue.ToString()))
                    {
                        ret.Add(o.Element);
                    }
                }
            }
            return ret;
        }

        private bool isStruct(string currentType)
        {
            bool ret = false;
            if (catalogtypes.Contains(currentType))
            {
                ret = true;
            }
            return ret;
        }
        
        private static void initCatalogTypes()
        {
            StreamReader sr = new StreamReader(catalogPath);
            string line;
            while((line=sr.ReadLine()) != null)
            {
                catalogtypes.Add(line);
            }
        }


        /* Dictionary<string, Dictionary<string, string>>
         * string:元素ID
         * Dictionary<string, string>:从getElementGeoProperties得到元素属性字典
         */
        public Dictionary<string, Dictionary<string, string>> getElementsGeoProperties(HashSet<IModelElement> eles)
        {
            Dictionary<string, Dictionary<string, string>> ret = new Dictionary<string, Dictionary<string, string>>();
            foreach (IModelElement ele in eles)
            {
                ret.Add(ele.Element.ElementId.ToString(), getElementGeoProperties(ele));
            }
            return ret;
        }

        private Dictionary<string, string> getElementGeoProperties(IModelElement ele)
        {
            Dictionary<string, string> geoProperties = new Dictionary<string, string>();
            geoProperties.Add("ID", ele.Element.ElementId.ToString());
            foreach (Dynamics o in ele.Objects)
            {
                IEnumerable<string> eleProps = o.GetDynamicMemberNames();
                IEnumerable<string> intersectprops = eleProps.Intersect(tartgetProps);

                foreach (string prop in intersectprops)
                {
                    geoProperties.Add(prop, o.ECInstance[prop].NativeValue.ToString());
                }

            }
            return geoProperties;
        }
    }
}
