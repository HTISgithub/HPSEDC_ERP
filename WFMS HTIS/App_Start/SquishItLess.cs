[assembly: WebActivator.PreApplicationStartMethod(typeof(Payroll.portal.App_Start.SquishItLess), "Start")]

namespace Payroll.portal.App_Start
{
    using SquishIt.Framework;
    using SquishIt.Less;

    public class SquishItLess
    {
        public static void Start()
        {
            Bundle.RegisterStylePreprocessor(new LessPreprocessor());
        }
    }
}