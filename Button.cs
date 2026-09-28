using System;
using System.Runtime.InteropServices;
using Inventor;
using System.Text;
using System.Drawing;
//using Microsoft.VisualBasic.Compatibility.VB6;
using System.Windows.Forms;
using InvDoc;

namespace InvAddIn
{                        
    internal abstract class Button
    {
#region Data Members
        private static Inventor.Application m_inventorApplication;
        private ButtonDefinition m_buttonDefinition;
        private ButtonDefinitionSink_OnExecuteEventHandler ButtonDefinition_OnExecuteEventDelegate;
#endregion
#region "Properties"
        public static Inventor.Application InventorApplication
        {
            set            
            {
                m_inventorApplication = value;
            }
            get
            {
                return m_inventorApplication;
            }
        }
        public Inventor.ButtonDefinition _ButtonDefinition
        {
            get
            {                                       
                return m_buttonDefinition;
            }
        }
#endregion
#region "Methods"
        public Button(string DisplayName, string internalName, string clientId, string description, string tooltip, Icon standarticon, Icon largeIcon, string catName = "CSharp")
        {
            try
            {
                stdole.IPictureDisp standartIconIPictureDisp;
                standartIconIPictureDisp = OleCreateConverter.ImageToPictureDisp(standarticon.ToBitmap());
                stdole.IPictureDisp largeIconIPictureDisp;
                largeIconIPictureDisp = OleCreateConverter.ImageToPictureDisp(largeIcon.ToBitmap());
                m_buttonDefinition = m_inventorApplication.CommandManager.ControlDefinitions.AddButtonDefinition(DisplayName,
                    "Autodesk:Macros:" + internalName, CommandTypesEnum.kNonShapeEditCmdType, clientId, description, tooltip, standartIconIPictureDisp,
                    largeIconIPictureDisp, ButtonDisplayEnum.kDisplayTextInLearningMode);
                m_buttonDefinition.Enabled = true; 
                ButtonDefinition_OnExecuteEventDelegate = new ButtonDefinitionSink_OnExecuteEventHandler(ButtonDefinition_OnExecute);
                m_buttonDefinition.OnExecute += ButtonDefinition_OnExecuteEventDelegate;
                if (catName != "")
                {
                    CommandCategory cat = u.get<CommandCategory>(m_inventorApplication.CommandManager.CommandCategories, c => c.DisplayName == catName) ??
                        m_inventorApplication.CommandManager.CommandCategories.Add(catName, "Autodesk:Macros:" + catName);
                    cat.Add(m_buttonDefinition);
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public Button(string displayName, string internalName, CommandTypesEnum commandType, string clientId, string description, string tooltip, ButtonDisplayEnum buttonDisplayType)
        {
            try
            {
                m_buttonDefinition = m_inventorApplication.CommandManager.ControlDefinitions.AddButtonDefinition(displayName, "Autodesk:Macros:" + internalName,
                    commandType, clientId, description, tooltip,
                    Type.Missing,Type.Missing,buttonDisplayType);
                m_buttonDefinition.Enabled = true;
                ButtonDefinition_OnExecuteEventDelegate = new ButtonDefinitionSink_OnExecuteEventHandler(ButtonDefinition_OnExecute);
                m_buttonDefinition.OnExecute += ButtonDefinition_OnExecuteEventDelegate;
            }
            catch (System.Exception ex)
            {
	            MessageBox.Show(ex.ToString());
            }
        }

        // горячая клавиша обязательна: если сочетание занято другой командой (E_FAIL), отбираем его у неё и назначаем повторно
        public void addShortCut(string s)
        {
            if (m_buttonDefinition.IsShortcutOverridden && m_buttonDefinition.OverrideShortcut == s) return;
            try { m_buttonDefinition.OverrideShortcut = s; return; }
            catch (System.Runtime.InteropServices.COMException) { }

            var owners = new System.Collections.Generic.List<string>();
            foreach (ControlDefinition cd in InventorApplication.CommandManager.ControlDefinitions)
            {
                if (cd.InternalName == m_buttonDefinition.InternalName) continue;
                string ov = "", def = "";
                try { if (cd.IsShortcutOverridden) ov = cd.OverrideShortcut; } catch { }
                try { def = cd.DefaultShortcut; } catch { }
                if (!string.Equals(ov, s, StringComparison.OrdinalIgnoreCase) && !string.Equals(def, s, StringComparison.OrdinalIgnoreCase)) continue;
                owners.Add(cd.InternalName + " (" + cd.DisplayName + ")");
                try { cd.OverrideShortcut = ""; } catch { }
            }
            try { m_buttonDefinition.OverrideShortcut = s; }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                System.Windows.Forms.MessageBox.Show("Не удалось назначить горячую клавишу \"" + s + "\" команде \"" + m_buttonDefinition.DisplayName + "\".\n" +
                    (owners.Count > 0 ? "Клавиша занята: " + string.Join(", ", owners) : "Занявшая её команда не найдена.") +
                    "\n\n" + ex.Message, "Macros");
            }
        }

        abstract protected void ButtonDefinition_OnExecute(NameValueMap context);

#endregion
    }
// 
//     internal class AxHostConverter : AxHost
// 
// {
// 
//     private AxHostConverter() : base("") { }
// 
//  
// 
//     static public stdole.IPictureDisp ImageToPictureDisp(Image image)
// 
//     {
// 
//         return (stdole.IPictureDisp)GetIPictureDispFromPicture(image);
// 
//     }
// 
//  
// 
//     static public Image PictureDispToImage(stdole.IPictureDisp pictureDisp)
// 
//     {
// 
//         return GetPictureFromIPicture(pictureDisp);
// 
//     }
// 
// }

    internal class OleCreateConverter

{

    [DllImport("oleaut32.dll", EntryPoint = "OleCreatePictureIndirect",

        CharSet = CharSet.Ansi, ExactSpelling = true, PreserveSig = true)]

    private static extern int OleCreatePictureIndirect(

        [In] PictDescBitmap pictdesc, ref Guid iid, bool fOwn,

        [MarshalAs(UnmanagedType.Interface)] out object ppVoid);

 

    const short _PictureTypeBitmap = 1;

    [StructLayout(LayoutKind.Sequential)]

    internal class PictDescBitmap

    {

        internal int cbSizeOfStruct = Marshal.SizeOf(typeof(PictDescBitmap));

        internal int pictureType = _PictureTypeBitmap;

        internal IntPtr hBitmap = IntPtr.Zero;

        internal IntPtr hPalette = IntPtr.Zero;

        internal int unused = 0;

 

        internal PictDescBitmap(Bitmap bitmap)

        {

            this.hBitmap = bitmap.GetHbitmap();

        }

    }

 

    public static stdole.IPictureDisp ImageToPictureDisp(Image image)

    {

        if (image == null || !(image is Bitmap))

        {

            return null;

        }

 

        PictDescBitmap pictDescBitmap = new PictDescBitmap((Bitmap)image);

        object ppVoid = null;

        Guid iPictureDispGuid = typeof(stdole.IPictureDisp).GUID;

        OleCreatePictureIndirect(pictDescBitmap, ref iPictureDispGuid, true, out ppVoid);

        stdole.IPictureDisp picture = (stdole.IPictureDisp)ppVoid;

        return picture;

    }

 

    public static Image PictureDispToImage(stdole.IPictureDisp pictureDisp)

    {

        Image image = null;

        if (pictureDisp != null && pictureDisp.Type == _PictureTypeBitmap)

        {

            IntPtr paletteHandle = new IntPtr(pictureDisp.hPal);

            IntPtr bitmapHandle = new IntPtr(pictureDisp.Handle);

            image = Image.FromHbitmap(bitmapHandle, paletteHandle);

        }

        return image;

    }

}

    public abstract class InvComboBox
    {
        public static Inventor.Application m_inventorApplication;
        public ComboBoxDefinition m_ComboBoxDef;
        public ComboBoxDefinitionSink_OnSelectEventHandler m_cbEhDelegate;
        public static Inventor.Application InventorApplication
        {
            set
            {
                m_inventorApplication = value;
            }
            get
            {
                return m_inventorApplication;
            }
        }
        public Inventor.ComboBoxDefinition _ComboBoxDef
        {
            get
            {
                return m_ComboBoxDef;
            }
        }
        public InvComboBox(string DisplayName, string InternalName, CommandTypesEnum Classification, int DropDownWidth,
            object ClientId, object StandardIcon,object LargeIcon, string DescriptionText = "", string ToolTipText = "", 
             ButtonDisplayEnum ButtonDisplay = ButtonDisplayEnum.kDisplayTextInLearningMode)
        {
            m_ComboBoxDef = m_inventorApplication.CommandManager.ControlDefinitions.AddComboBoxDefinition(DisplayName, InternalName, Classification, DropDownWidth, ClientId,
                DescriptionText, ToolTipText, StandardIcon, LargeIcon);
            m_ComboBoxDef.Enabled = true;
            m_cbEhDelegate = new ComboBoxDefinitionSink_OnSelectEventHandler(ComboBoxDefinition_OnSelect);
            m_ComboBoxDef.OnSelect += m_cbEhDelegate;
        }

        public InvComboBox(string DisplayName, string InternaleName, CommandTypesEnum Class, int Drop)
        {
            m_ComboBoxDef = m_inventorApplication.CommandManager.ControlDefinitions.AddComboBoxDefinition(DisplayName, InternaleName, Class, Drop);
            m_ComboBoxDef.Enabled = true;
            m_cbEhDelegate = new ComboBoxDefinitionSink_OnSelectEventHandler(ComboBoxDefinition_OnSelect);
            m_ComboBoxDef.OnSelect += m_cbEhDelegate;
        }

        abstract protected void ComboBoxDefinition_OnSelect(NameValueMap nvm);
    }
}



