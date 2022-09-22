using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace doc2docxByWps
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // 需要安装wps
            // 引用 Upgrade Kingsoft WPS 3.0 Object Library
            if(args.Length != 2)
            {
                Console.WriteLine("参数有误");
            }
            string srcPath=args[0];
            string destPath=args[1];
            if (System.IO.File.Exists(srcPath))
            {
                var type = Type.GetTypeFromProgID("KWps.Application");
                dynamic wps = Activator.CreateInstance(type);
                try
                {
                    dynamic doc = wps.Documents.Open(srcPath, Visible: false);
                    doc.SaveAs(destPath, WdSaveFormat.wdFormatXMLDocument);
                    doc.close();
                    Console.WriteLine(destPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"转换docx出错:{ex.StackTrace}");
                }
                finally
                {
                    wps.Quit();
                }
            }
            else
            {
                Console.WriteLine("源文件不存在");
            }
        }
    }
}
