namespace BasicInfo
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var js = new BasicInfo();
            var rd = js.ReadDirectory();
            var tj = js.TransformToString(rd);
            var sj = js.GetMapExcel(tj);

            //var st = new StateTax();
            //var stRd = st.ReadDirectory();
            //var stTs = st.TransformToString(stRd);
            //var stSte= st.StateTaxExcel(stTs);

            //var tr = new TaxRegime();
            //var trRd = tr.ReadDirectory();
            //var trTs = tr.TransformToString(trRd);
            //var trTre = tr.TaxRegimeExcel(trTs);

        }
    }
}
