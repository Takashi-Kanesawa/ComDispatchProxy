using Excel = Microsoft.Office.Interop.Excel;

namespace ComDispatchProxy.ProxyFactories;

public class InteropExcelProxyFactory : ComProxyFactoryBase
{
    protected sealed override string TargetAssemblyName => "Microsoft.Office.Interop.Excel";

    public InteropExcelProxyFactory(IEnumerable<Type> usedTypes) : base(usedTypes)
    {
    }

    protected override IEnumerable<Type> FilteredTypes =>
        base.FilteredTypes.Where(t =>
            t.IsInterface &&
            (t.Name.StartsWith("_", StringComparison.Ordinal)) == false);
}
