import os.path
from cffi import FFI

src_dir = os.path.join(os.path.dirname(__file__), "src")

cdef = open(os.path.join(src_dir, "XLCALLAPI.H")).read()

ffi = FFI()
ffi.cdef(
    cdef
    + """
extern int Excel12(int xlfn, LPXLOPER12 operRes, int count, ... );
// wrapped xloper which includes a reference to tself
"""
)

ffi.set_source(
    "_xlcall",
    r"""
#include <WINDOWS.H>
#include <XLCALL.H>

#define cxloper12Max 255
#define EXCEL12ENTRYPT "MdCallBack12"

typedef int (PASCAL *EXCEL12PROC) (
    int xlfn, 
    int coper, 
    LPXLOPER12 *rgpxloper12, 
    LPXLOPER12 xloper12Res
);
HMODULE hmodule;
EXCEL12PROC pexcel12;

__forceinline void FetchExcel12EntryPt(void)
{
   if (pexcel12 == NULL)
   {
       hmodule = GetModuleHandle(NULL);
       if (hmodule != NULL)
       {
           pexcel12 = (EXCEL12PROC) GetProcAddress(hmodule, EXCEL12ENTRYPT);
       }
   }
}

void pascal SetExcel12EntryPt(EXCEL12PROC pexcel12New)
{
   FetchExcel12EntryPt();
   if (pexcel12 == NULL)
   {
       pexcel12 = pexcel12New;
   }
}

int __cdecl Excel12(int xlfn, LPXLOPER12 operRes, int count, ...)
{
   LPXLOPER12 rgxloper12[cxloper12Max];
   va_list ap;
   int ioper;
   int mdRet;

   FetchExcel12EntryPt();
   if (pexcel12 == NULL)
   {
       OutputDebugString("No Excel12EntryPt found!\n");
       mdRet = xlretFailed;
   }
   else
   {
       mdRet = xlretInvCount;
       if ((count >= 0)  && (count <= cxloper12Max))
       {
           va_start(ap, count);
           for (ioper = 0; ioper < count ; ioper++)
           {
               rgxloper12[ioper] = va_arg(ap, LPXLOPER12);
           }
           va_end(ap);			
           mdRet = (pexcel12)(xlfn, count, &rgxloper12[0], operRes);
       }
   }
   return(mdRet);
}""",
    include_dirs=[src_dir],
)

if __name__ == "__main__":
    ffi.compile(target=os.path.join(os.path.dirname(__file__), "_xlcall.pyd"))
