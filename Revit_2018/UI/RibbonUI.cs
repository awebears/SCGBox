using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;

namespace Revit_2018.UI
{
    [Transaction(TransactionMode.Manual)]
    internal class RibbonUI : IExternalApplication
    {
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            #region 创建RibbonTab：SCGBox
            string tabName = "SCGBox";
            application.CreateRibbonTab(tabName);
            #endregion

            string assemeblyLocation = GetType().Assembly.Location;

            #region 创建Panel：基准
            string panelName_Standardize = "基准";
            RibbonPanel ribbonPanel_Standardize = application.CreateRibbonPanel(tabName, panelName_Standardize);
            #region 添加按钮：标高
            string className_Standardize = "Revit_2018.ExcutionLibrary.Datum.StandardizeLevel";
            PushButtonData pushButtonData_Standardize = new PushButtonData("internalName_StanderdardizeLevel", "标高", assemeblyLocation, className_Standardize)
            { LargeImage = SetIcon("Revit_2018.UI.icon.标高_32px.png"), Image = SetIcon("Revit_2018.UI.icon.标高_16px.png"), ToolTip = "修改标高名称，统一命名方式" };
            ribbonPanel_Standardize.AddItem(pushButtonData_Standardize);
            #endregion
            #endregion

            #region 创建Panel：建筑
            #endregion

            #region 创建Panel：结构
            #endregion

            #region 创建Panel：机电
            string panelName_MEP = "机电";
            RibbonPanel ribbonPanel_MEP = application.CreateRibbonPanel(tabName, panelName_MEP);

            string className_SwitchOn = "Revit_2018.ExcutionLibrary.MEP.SwitchOn";
            PushButtonData pushButtonData_SwitchOn = new PushButtonData("internal_SwitchOn", "全部打开", assemeblyLocation, className_SwitchOn)
            { LargeImage = SetIcon("Revit_2018.UI.icon.打开_32px.png"), Image = SetIcon("Revit_2018.UI.icon.打开_16px.png"), ToolTip = "打开当前视图的所有过滤器" };

            string className_SwitchOff = "Revit_2018.ExcutionLibrary.MEP.SwitchOff";
            PushButtonData pushButtonData_SwitchOff = new PushButtonData("internal_SwitchOff", "全部关闭", assemeblyLocation, className_SwitchOff)
            { LargeImage = SetIcon("Revit_2018.UI.icon.关闭_32px.png"), Image = SetIcon("Revit_2018.UI.icon.关闭_16px.png"), ToolTip = "关闭当前视图的所有过滤器" };

            string className_SwitchSelectedOn = "Revit_2018.ExcutionLibrary.MEP.SwitchSelectedOn";
            PushButtonData pushButtonData_SwitchSelectedOn = new PushButtonData("internal_SwitchSelectedOn", "隔离选择", assemeblyLocation, className_SwitchSelectedOn)
            { LargeImage = SetIcon("Revit_2018.UI.icon.隔离_32px.png"), Image = SetIcon("Revit_2018.UI.icon.隔离_16px.png"), ToolTip = "隔离选中图元所属的过滤器" };

            string className_SwitchSelectedOff = "Revit_2018.ExcutionLibrary.MEP.SwitchSelectedOff";
            PushButtonData pushButtonData_SwitchSelectedOff = new PushButtonData("internal_SwitchSelectedOff", "选择关闭", assemeblyLocation, className_SwitchSelectedOff)
            { LargeImage = SetIcon("Revit_2018.UI.icon.点击_32px.png"), Image = SetIcon("Revit_2018.UI.icon.点击_16px.png"), ToolTip = "关闭选中图元所属的过滤器" };

            string className_CutMepByLine = "Revit_2018.ExcutionLibrary.MEP.CutAllMepByLine";
            PushButtonData pushButtonData_CutMepByLine = new PushButtonData("internal_CutAllMepByLine", "一键打断", assemeblyLocation, className_CutMepByLine)
            { LargeImage = SetIcon("Revit_2018.UI.icon.断开_32px.png"), Image = SetIcon("Revit_2018.UI.icon.断开_16px.png"), ToolTip = "用一条线打断所有水管、风管、桥架" };

            ribbonPanel_MEP.AddItem(pushButtonData_SwitchOn);
            ribbonPanel_MEP.AddItem(pushButtonData_SwitchOff);
            ribbonPanel_MEP.AddItem(pushButtonData_SwitchSelectedOn);
            ribbonPanel_MEP.AddItem(pushButtonData_SwitchSelectedOff);

            ribbonPanel_MEP.AddSeparator();

            ribbonPanel_MEP.AddItem(pushButtonData_CutMepByLine);
            #endregion  

            #region 创建Panel：通用
            string panelName_General = "通用";
            RibbonPanel ribbonPanel_General = application.CreateRibbonPanel(tabName, panelName_General);

            #region 链接功能区
            RadioButtonGroupData radioButtonGroupData_Link = new RadioButtonGroupData("internal_Link_Locates");
            RadioButtonGroup radioButtonGroup_Link = ribbonPanel_General.AddItem(radioButtonGroupData_Link) as RadioButtonGroup;
            radioButtonGroup_Link.ItemText = "定位方式";
            //togglebutton
            ToggleButtonData toggleButtonDara_Link_Origin = new ToggleButtonData("Origin", "原点")
            { LargeImage = SetIcon("Revit_2018.UI.icon.原点_32px.png"), Image = SetIcon("Revit_2018.UI.icon.原点_16px.png"), ToolTip = "定位方式：原点到原点" };
            
            ToggleButtonData toggleButtonDara_Link_Shared = new ToggleButtonData("Shared", "共享坐标")
            { LargeImage = SetIcon("Revit_2018.UI.icon.共享坐标_32px.png"), Image = SetIcon("Revit_2018.UI.icon.共享坐标_16px.png"), ToolTip = "定位方式：共享坐标" };
            radioButtonGroup_Link.AddItem(toggleButtonDara_Link_Origin);
            radioButtonGroup_Link.AddItem(toggleButtonDara_Link_Shared);

            //添加分隔符
            ribbonPanel_General.AddSeparator();
            //pushbutton
            string className_Link_Load = "Revit_2018.ExcutionLibrary.General.LoadLinks";
            PushButtonData pushButtonData_Link_Load = new PushButtonData("internal_LoadLink", "加载", assemeblyLocation, className_Link_Load)
            { LargeImage = SetIcon("Revit_2018.UI.icon.加载_32px.png"), Image = SetIcon("Revit_2018.UI.icon.加载_16px.png"), ToolTip = "批量插入链接模型，先选择定位方式" };

            string className_Link_Reload = "Revit_2018.ExcutionLibrary.General.ReloadLinks";
            PushButtonData pushButtonData_Link_Reload = new PushButtonData("internal_ReloadLink", "重载", assemeblyLocation, className_Link_Reload)
            { LargeImage = SetIcon("Revit_2018.UI.icon.重载_32px.png"), Image = SetIcon("Revit_2018.UI.icon.重载_16px.png"), ToolTip = "重新载入所有链接模型，先选择定位方式" };

            string className_Link_Unload = "Revit_2018.ExcutionLibrary.General.UnloadLinks";
            PushButtonData pushButtonData_Link_Unload = new PushButtonData("internal_UnloadLink", "卸载", assemeblyLocation, className_Link_Unload)
            { LargeImage = SetIcon("Revit_2018.UI.icon.卸载_32px.png"), Image = SetIcon("Revit_2018.UI.icon.卸载_16px.png"), ToolTip = "卸载所有链接模型" };

            string className_Link_Delete = "Revit_2018.ExcutionLibrary.General.DeleteLinks";
            PushButtonData pushButtonData_Link_Delete = new PushButtonData("internal_DeleteLink", "删除", assemeblyLocation, className_Link_Delete)
            { LargeImage = SetIcon("Revit_2018.UI.icon.删除_32px.png"), Image = SetIcon("Revit_2018.UI.icon.删除_16px.png"), ToolTip = "删除所有链接模型" };

            ribbonPanel_General.AddStackedItems(pushButtonData_Link_Load, pushButtonData_Link_Reload);
            ribbonPanel_General.AddStackedItems(pushButtonData_Link_Unload, pushButtonData_Link_Delete);
            #endregion

            #region 剪切功能区
            ribbonPanel_General.AddSeparator();
            string className_CorrectJoinRelationships = "Revit_2018.ExcutionLibrary.General.CorrectJoinRelationships";
            PushButtonData pushButtonData_CorrectJoinRelationships = new PushButtonData("internalName_CorrectJoinRelationships ", "切换连接", assemeblyLocation, className_CorrectJoinRelationships)
            { LargeImage = SetIcon("Revit_2018.UI.icon.切换_32px.png"), Image = SetIcon("Revit_2018.UI.icon.切换_16px.png"), ToolTip = "调整建筑结构构件剪切关系" };
            ribbonPanel_General.AddItem(pushButtonData_CorrectJoinRelationships);
            #endregion
            #endregion

            return Result.Succeeded;
        }
        private BitmapImage SetIcon(string iconPath)
        {
            Stream stream = this.GetType().Assembly.GetManifestResourceStream(iconPath);
            BitmapImage icon = new BitmapImage();
            icon.BeginInit();
            icon.StreamSource = stream;
            icon.EndInit();
            return icon;
        }
    }
}
