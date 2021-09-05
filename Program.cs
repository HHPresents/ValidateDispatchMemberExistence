using System;
using System.Runtime.InteropServices;

namespace ValidateDispatchMemberExistence
{
    class Program
    {
        static void Main(string[] args)
        {
            object comObject = Activator.CreateInstance(Type.GetTypeFromProgID("Scripting.FileSystemObject"));

            string validateMemberName;
            bool result;

            validateMemberName = "CreateTextFile";
            result = DispatchMemberValidator.Exists(comObject, validateMemberName);
            Console.WriteLine("Scripting.FileSystemObject に {0} メソッドもしくはプロパティ は {1}。", validateMemberName, result ? "存在する" : "存在しない");

            validateMemberName = "CreateBinaryFile";
            result = DispatchMemberValidator.Exists(comObject, validateMemberName);
            Console.WriteLine("Scripting.FileSystemObject に {0} メソッドもしくはプロパティ は {1}。", validateMemberName, result ? "存在する" : "存在しない");

            validateMemberName = "Drives";
            result = DispatchMemberValidator.Exists(comObject, validateMemberName);
            Console.WriteLine("Scripting.FileSystemObject に {0} メソッドもしくはプロパティ は {1}。", validateMemberName, result ? "存在する" : "存在しない");

            Console.WriteLine("Hit Any Key.");
            Console.ReadKey();
        }
    }

    static class DispatchMemberValidator
    {
        /// <summary>
        /// COMオブジェクトの中にメンバー(メソッド/プロパティ)は存在するか？
        /// </summary>
        /// <param name="comObject">検証対象のCOMオブジェクト。IDispatchが実装されている前提。</param>
        /// <param name="memberName">確認したいメンバー名</param>
        /// <returns>存在する場合はtrue。存在しない場合はfalse。検証時にエラーが発生した場合はComExceptionをスローする。</returns>
        public static bool Exists(object comObject, string memberName)
        {
            Guid iidNull = new Guid();
            var dispatch = (IDispatchCOM)comObject;
            int dispid;
            int hr = dispatch.GetIDsOfNames(ref iidNull, ref memberName, 1, LOCALE_USER_DEFAULT, out dispid);

            if (hr == 0)
            {
                return true;
            }
            else if (hr == DISP_E_UNKNOWNNAME)
            {
                return false;
            }
            else
            {
                Marshal.ThrowExceptionForHR(hr);
                return false; // ※ここに到達しない事は明白だが書かないとコンパイルに失敗する・・・
            }
        }
        
        private const int LOCALE_USER_DEFAULT = 1024;
        private const int DISP_E_UNKNOWNNAME = unchecked((int)0x80020006);

        /// <summary>
        /// IDispatchのなんちゃって定義(GetIDsOfNames以外はデタラメ)
        /// </summary>
        [Guid("00020400-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IDispatchCOM
        {
            // 使用しないのでシグネチャはデタラメ。ただし定義順序(vtbl)は重要なので移動させたり消してはいけない。
            [PreserveSig]
            int GetTypeInfoCount();

            // 使用しないのでシグネチャはデタラメ。ただし定義順序(vtbl)は重要なので移動させたり消してはいけない。
            [PreserveSig]
            int GetTypeInfo();

            // 使用するのでシグネチャは正確。定義順序(vtbl)は重要なので移動させてはいけない。
            [PreserveSig]
            int GetIDsOfNames(ref Guid riid, [MarshalAs(UnmanagedType.LPWStr)] ref string rgszNames, int cNames, int lcid, out int rgDispId);

            // 使用しないのでシグネチャはデタラメ。ただし定義順序(vtbl)は重要なので移動させたり消してはいけない。
            [PreserveSig]
            int Invoke();
        }

    }
}
