using Excel = Microsoft.Office.Interop.Excel;

namespace ComDispatchProxy.ProxyFactories;

public class InteropExcelProxyFactory : ComProxyFactoryBase
{
    protected sealed override string TargetAssemblyName => "Microsoft.Office.Interop.Excel";

    public InteropExcelProxyFactory(string xmlPath) : base(xmlPath)
    {
    }

    protected override IEnumerable<Type> GetUsedTypes()
    {
        return typeof(Excel.Application).Assembly.GetTypes()
            .Where(t => t.IsInterface && t.Namespace == TargetAssemblyName);
    }
}
