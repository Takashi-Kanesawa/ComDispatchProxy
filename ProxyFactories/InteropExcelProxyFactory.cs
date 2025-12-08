using Excel = Microsoft.Office.Interop.Excel;

namespace ComDispatchProxy.ProxyFactories;

public class InteropExcelProxyFactory : ComProxyFactoryBase
{
    protected sealed override string TargetAssemblyName => "Microsoft.Office.Interop.Excel";

    public InteropExcelProxyFactory(string xmlPath, IEnumerable<Type> usedTypes) : base(xmlPath, usedTypes)
    {
    }
}
