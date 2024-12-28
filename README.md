# VBScript Late Binding Bug

This repository demonstrates a common issue in VBScript programming: runtime errors caused by late binding.  Late binding, while offering flexibility, can lead to unexpected failures if the referenced object or method is unavailable at runtime. The provided example showcases this issue and offers a solution using early binding for improved robustness.

## Bug Description

The `lateBindingBug.vbs` script attempts to create an Excel application object and display its version.  If Microsoft Excel is not installed on the system, the script will fail with a runtime error because the `CreateObject` call will not succeed.  This is characteristic of late binding's lack of compile-time type checking.

## Solution

The `lateBindingSolution.vbs` script demonstrates a solution by using early binding. While it requires declaring the type, this provides compile-time verification and will prevent runtime failures related to object existence.

## How to Reproduce

1.  Save `lateBindingBug.vbs`.
2.  Run the script.  If Excel is installed, it'll work; otherwise, a runtime error will occur.
3.  Save `lateBindingSolution.vbs`. This uses early binding to address the issue. It will provide a compile-time or runtime error, depending on the VBScript interpreter used and the presence of Excel, but the error is more predictable and informative.