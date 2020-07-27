using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraGrid.Views.Grid;

namespace LabelPrinttingSystem
{
    public partial class ShipConfirm : DevExpress.XtraEditors.XtraForm
    {
        private DataTable _Parameter;
        SizeF TSize;

        public ShipConfirm(DataTable parameter)
        {
            InitializeComponent();

            _Parameter = parameter;
        }

        private void ShipConfirm_Load(object sender, EventArgs e)
        {
            grdShipConfirm.DataSource = _Parameter;
            grdShipConfirm.RefreshDataSource();
            
            this.grvShipConfirm.CustomDrawRowIndicator += grvShipConfirm_CustomDrawRowIndicator;
        }
        
        /// <summary>
        /// Indicator에 Row Num 자동 채번
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void grvShipConfirm_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            DevExpress.XtraGrid.Views.Grid.GridView view = sender as DevExpress.XtraGrid.Views.Grid.GridView;
            //Seq 작업
            if (e.Info.IsRowIndicator && e.RowHandle >= 0)
            {
                //우측정렬
                e.Info.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
                //0부터 되기 때문에 +1 진행
                e.Info.DisplayText = (e.RowHandle + 1).ToString();

                TSize = e.Graphics.MeasureString(e.Info.DisplayText, e.Appearance.Font);
                int requiredWidth = Convert.ToInt32(TSize.Width) + 16;
                view.IndicatorWidth = view.IndicatorWidth < requiredWidth ? requiredWidth : view.IndicatorWidth;
            }
        }
    }
}
