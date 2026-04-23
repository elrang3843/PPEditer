# Third-Party Notices — PPEditer

PPEditer includes the following third-party open-source components.
All are distributed under the **MIT License**, which is reproduced in full for each component below.

---

## 1. .NET 8 Runtime & WPF

**Copyright © .NET Foundation and Contributors**  
**Copyright © Microsoft Corporation**

- Repository: https://github.com/dotnet/wpf  
- NuGet: https://www.nuget.org/packages/Microsoft.WindowsDesktop.App.WPF

> MIT License
>
> Permission is hereby granted, free of charge, to any person obtaining a copy
> of this software and associated documentation files (the "Software"), to deal
> in the Software without restriction, including without limitation the rights
> to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
> copies of the Software, and to permit persons to whom the Software is
> furnished to do so, subject to the following conditions:
>
> The above copyright notice and this permission notice shall be included in all
> copies or substantial portions of the Software.
>
> THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
> IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
> FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
> AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
> LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
> OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
> SOFTWARE.

> **Patent Notice**: Microsoft provides additional patent grants for .NET under a separate
> patent promise. See https://github.com/dotnet/runtime/blob/main/PATENTS.TXT

---

## 2. DocumentFormat.OpenXml SDK v3.2.0

**Copyright © Microsoft Corporation**

- Repository: https://github.com/dotnet/Open-XML-SDK  
- NuGet: https://www.nuget.org/packages/DocumentFormat.OpenXml

> MIT License
>
> Permission is hereby granted, free of charge, to any person obtaining a copy
> of this software and associated documentation files (the "Software"), to deal
> in the Software without restriction, including without limitation the rights
> to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
> copies of the Software, and to permit persons to whom the Software is
> furnished to do so, subject to the following conditions:
>
> The above copyright notice and this permission notice shall be included in all
> copies or substantial portions of the Software.
>
> THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
> IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
> FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
> AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
> LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
> OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
> SOFTWARE.

> **OOXML Format Notice**: Reading and writing .pptx / .docx / .xlsx files relies on the
> Office Open XML (OOXML / Ecma-376 / ISO/IEC 29500) specification.
> Microsoft's Open Specification Promise (OSP) grants royalty-free rights to implement
> this specification. See https://docs.microsoft.com/openspecs/dev_center/ms-devcentlp

---

## 3. PDFsharp v6.1.1

**Copyright © empira Software GmbH, Cologne Area (Germany)**

- Repository: https://github.com/empira/PDFsharp  
- NuGet: https://www.nuget.org/packages/PDFsharp

> MIT License
>
> Copyright (c) empira Software GmbH, Cologne Area (Germany)
>
> Permission is hereby granted, free of charge, to any person obtaining a copy
> of this software and associated documentation files (the "Software"), to deal
> in the Software without restriction, including without limitation the rights
> to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
> copies of the Software, and to permit persons to whom the Software is
> furnished to do so, subject to the following conditions:
>
> The above copyright notice and this permission notice shall be included in all
> copies or substantial portions of the Software.
>
> THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
> IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
> FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
> AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
> LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
> OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
> SOFTWARE.

---

## 4. System.IO.Packaging (transitive)

**Copyright © .NET Foundation and Contributors**

- Repository: https://github.com/dotnet/runtime  
- NuGet: https://www.nuget.org/packages/System.IO.Packaging

Distributed under the MIT License. See component 1 for the full license text.

---

## 5. Native WPF Runtime DLLs

The following native (non-managed) DLL files are included in the self-contained deployment.
They are produced by Microsoft and distributed under **Microsoft's .NET Redistribution Rights**,
which explicitly permit redistribution as part of self-contained .NET applications.

| File | Purpose |
|------|---------|
| `coreclr.dll` | .NET CoreCLR runtime host |
| `clrjit.dll` | .NET JIT compiler |
| `clrgc.dll` | .NET garbage collector |
| `wpfgfx_cor3.dll` | WPF graphics engine |
| `PresentationNative_cor3.dll` | WPF native bridge |
| `PenImc_cor3.dll` | WPF pen / touch input |
| `vcruntime140_cor3.dll` | Visual C++ runtime for .NET |
| `D3DCompiler_47.dll` | DirectX shader compiler (WPF hardware rendering) |

**Copyright © Microsoft Corporation**

Redistribution of these files is permitted under:
- **.NET Redistribution Policy**: https://github.com/dotnet/core/blob/main/license-information.md  
- **Visual C++ Redistributable License** (for `vcruntime140_cor3.dll`): included in the Microsoft Visual C++ Redistributable package  
- **DirectX Redistribution Rights** (for `D3DCompiler_47.dll`): governed by the Windows SDK / DirectX REDIST terms; see `REDIST.TXT` in the Windows SDK

These files are provided **as-is** by Microsoft and are not modified by this project.  
Source for the MIT-licensed portions: https://github.com/dotnet/wpf  
Source for the .NET runtime: https://github.com/dotnet/runtime

---

*This file is provided to satisfy the copyright-notice preservation requirement of the MIT License,
and to document Microsoft's redistribution rights for native runtime components.*  
*The PPEditer application itself is also distributed under the MIT License — see [LICENSE](./LICENSE).*  
***This file must be included in every binary distribution (ZIP release) of PPEditer.***
