using System;
using System.Collections.Generic;
using System.Net.Sockets;
using System.Reflection;
using System.Text;

namespace Cqpaul.Dotnet.Util.Helpers
{
    public class ReflectFunction
    {
        /// <summary>
        /// 通过反射复制两个对象中，具有相同属性的值。 
        /// remainNdecimals:如果是double或者decimal类型的小数，则保留N位数，N默认为2位小数
        /// </summary>
        public static void CopyPropertys<Input, OutPut>(Input inputObj, OutPut outPutObj)
        {
            if (inputObj == null)
            {
                return; //为空则return
            }

            var Types = inputObj.GetType();//获得类型
            var Typed = typeof(OutPut);
            var validTypeList = new List<string>() { typeof(DateTime?).FullName, typeof(byte[]).FullName, typeof(DateTime).FullName, typeof(string).FullName, typeof(bool).FullName, typeof(bool?).FullName, typeof(long).FullName, typeof(long?).FullName, typeof(int).FullName, typeof(int?).FullName, typeof(double).FullName, typeof(double?).FullName, typeof(decimal).FullName, typeof(decimal?).FullName };
            foreach (PropertyInfo sp in Types.GetProperties())//获得类型的属性字段
            {
                if (validTypeList.Contains(sp.PropertyType.FullName))  //只有基础数据类型才转换。 eg:string ,boolean, long, in so on
                {
                    foreach (PropertyInfo dp in Typed.GetProperties())
                    {
                        if (dp.Name == sp.Name && dp.PropertyType.FullName == sp.PropertyType.FullName)//判断属性类型和名称相同的时候， 才赋值
                        {
                            var inputObjValue = sp.GetValue(inputObj, null);
                            dp.SetValue(outPutObj, inputObjValue, null);//获得s对象属性的值复制给d对象的属性
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="_object"></param>
        /// <returns></returns>
        public static object DeepCopy(object _object)
        {
            Type T = _object.GetType();
            object o = Activator.CreateInstance(T);
            PropertyInfo[] PI = T.GetProperties();
            for (int i = 0; i < PI.Length; i++)
            {
                PropertyInfo P = PI[i];
                P.SetValue(o, P.GetValue(_object));
            }
            return o;
        }


    }
}
