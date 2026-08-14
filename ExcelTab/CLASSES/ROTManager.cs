using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Excel = NetOffice.ExcelApi;

namespace ExcelTab.CLASSES
{
    public class ROTManager
    {
        [MethodImpl(MethodImplOptions.Synchronized)]
        public static List<Excel.Workbook> GetOpendLocalWorkbooks()
        {
            List<Excel.Workbook> rst = new List<Excel.Workbook>();
            foreach (var o in GetRunningObjects(MKSYS.MKSYS_FILEMONIKER))
            {
                Excel.Workbook wb = null;
                try
                {
                    wb = new Excel.Workbook(null, o);
                    if (IsIgnored(wb.Name))
                    {
                        wb.Dispose();
                    }
                    else
                    {
                        rst.Add(wb);
                    }
                }
                catch (Exception)
                {
                    if (wb != null)
                    {
                        ComHelper.ReleaseCom(wb);
                    }
                    else
                    {
                        ComHelper.ReleaseCom(o);
                    }
                }
            }
            return rst;
        }

        public static int GetOpenedLocalWorkbookCount()
        {
            List<Excel.Workbook> allWb = GetOpendLocalWorkbooks();
            int count = allWb.Count;
            ComHelper.ReleaseComList(allWb);
            return count;
        }

        private static bool IsIgnored(string name)
        {
            List<string> list = App.Setting?.IgnoreFiles;
            if (list == null)
            {
                return false;
            }
            for (int i = 0; i < list.Count; i++)
            {
                if (string.Equals(list[i], name, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// ファイルモニカを元に取得できるオブジェクトから、指定した型のオブジェクトを探す
        /// </summary>
        /// <typeparam name="T">取得したいCOMの型</typeparam>
        /// <returns>型変換に成功したオブジェクト</returns>
        public static T[] GetOpendComObjects<T>() where T : class
        {
            List<T> ts = new List<T>();
            // OfType<T>は解放のタイミングが難しそう
            foreach (var o in GetRunningObjects(MKSYS.MKSYS_FILEMONIKER))
            {
                try
                {
                    T t = o as T;
                    if (t != null)
                    {
                        ts.Add(t);
                    }
                    else
                    {
                        Marshal.FinalReleaseComObject(o);
                    }
                }
                catch(Exception)
                {
                    try
                    {
                        Marshal.FinalReleaseComObject(o);
                    }
                    catch (Exception) { }
                }
            }
            return ts.ToArray();
        }

        /// <summary>
        /// モニカーの種類
        /// </summary>
        /// <see cref="https://docs.microsoft.com/ja-jp/windows/desktop/api/objidl/ne-objidl-tagmksys"/>
        enum MKSYS
        {
            MKSYS_NONE,
            MKSYS_GENERICCOMPOSITE,
            MKSYS_FILEMONIKER,
            MKSYS_ANTIMONIKER,
            MKSYS_ITEMMONIKER,
            MKSYS_POINTERMONIKER,
            MKSYS_CLASSMONIKER,
            MKSYS_OBJREFMONIKER,
            MKSYS_SESSIONMONIKER,
            MKSYS_LUAMONIKER
        }

        /// <summary>
        /// Returns a pointer to an implementation of IBindCtx (a bind context object). This object stores information about a particular moniker-binding operation.
        /// </summary>
        /// <param name="reserved">This parameter is reserved and must be 0.</param>
        /// <param name="ppbc">Address of an IBindCtx* pointer variable that receives the interface pointer to the new bind context object. When the function is successful, the caller is responsible for calling Release on the bind context. A NULL value for the bind context indicates that an error occurred.</param>
        /// <returns>This function can return the standard return values E_OUTOFMEMORY and S_OK</returns>
        /// <see cref="https://docs.microsoft.com/en-us/windows/desktop/api/objbase/nf-objbase-createbindctx"/>
        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);

        /// <summary>
        /// Running Object Table から指定されたモニカのオブジェクトを取得する
        /// </summary>
        /// <param name="monikerType">モニカの種類</param>
        /// <returns>見つかったオブジェクト</returns>
        private static object[] GetRunningObjects(MKSYS monikerType)
        {
            List<object> returnObjects = new List<object>();
            try
            {
                // Running Object Table を取得する
                const uint reserved = 0;
                IBindCtx ctx;
                CreateBindCtx(reserved, out ctx);

                IRunningObjectTable runningObjectTable;
                ctx.GetRunningObjectTable(out runningObjectTable);

                IEnumMoniker enumMoniker;
                runningObjectTable.EnumRunning(out enumMoniker);
                // ここまではほぼ定型

                enumMoniker.Reset();

                while (true)
                {
                    const int S_OK = 0;

                    IMoniker[] tmpMks = new IMoniker[1];
                    IntPtr fetched = IntPtr.Zero;
                    // bufMonikers の数ずつモニカーを取得
                    bool successNext = enumMoniker.Next(tmpMks.Length, tmpMks, fetched) == S_OK;
                    if (!successNext)
                    {
                        break;
                    }

                    try
                    {
                        //for Debug
                        string dispName;
                        tmpMks[0].GetDisplayName(ctx, null, out dispName);
                        //Debug.WriteLine("DisplayName\t" + dispName);

                        Guid clsId;
                        tmpMks[0].GetClassID(out clsId);
                        //Debug.WriteLine(clsId);
                        // Debug.WriteLine(Marshal.GetTypeFromCLSID(clsId)); //-> System.__ComObject

                        int pdwMksys;
                        if (tmpMks[0].IsSystemMoniker(out pdwMksys) != S_OK)
                        {
                            Debug.WriteLine("not SystemMoniker");
                            continue;
                        }

                        MKSYS mkType = (MKSYS)Enum.ToObject(typeof(MKSYS), pdwMksys);
                        //Debug.WriteLine(mkType);
                        if (mkType != monikerType)
                        {
                            continue;
                        }

                        object obj;
                        if (runningObjectTable.GetObject(tmpMks[0], out obj) != S_OK)
                        {
                            continue;
                        }
                        returnObjects.Add(obj);
                    }
                    catch
                    {
                        //Debug.WriteLine(ex.Message);
                    }
                    finally
                    {
                        Marshal.FinalReleaseComObject(tmpMks[0]);
                    }
                }

                Marshal.FinalReleaseComObject(enumMoniker);
                Marshal.FinalReleaseComObject(runningObjectTable);
                Marshal.FinalReleaseComObject(ctx);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }

            return returnObjects.ToArray();
        }
    }
}
