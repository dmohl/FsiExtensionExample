module FsiExtension

open System.IO
open EnvDTE
open EnvDTE80
open System.Runtime.InteropServices
open NuGet

let packageSource = new PackageSource("https://go.microsoft.com/fwlink/?LinkID=206669")

let repositoryFactory = new PackageRepositoryFactory() 
let repository = repositoryFactory.CreateRepository (packageSource)

let automationObject = Marshal.GetActiveObject("VisualStudio.DTE.10.0")
let dte2 = automationObject :?> DTE2

let path = match dte2.Solution with
            | null -> dte2.ActiveDocument.Path
            | _ -> Path.GetDirectoryName dte2.Solution.FullName

let SendToFsiConsole text = 
    printfn "\r\n%s\r\n" text

let GetRepositoryPath path = 
    Path.Combine(path, "packages")

let InstallPackages (repositoryId:string) (destinationPath:string) =
    "Getting ready to retrieve the package. This may take several minutes..." |> SendToFsiConsole
    let packageManager = new PackageManager(repository, GetRepositoryPath destinationPath)
    packageManager.InstallPackage repositoryId

let AddReferenceToFsx referenceName =
    if dte2.ActiveDocument.Name.Contains ".fsx" then
        let doc = dte2.ActiveDocument.Object "TextDocument" :?> TextDocument
        let editPoint = doc.CreateEditPoint()
        if not ((editPoint.GetText doc.EndPoint).Contains(referenceName)) then
            editPoint.LineDown 1
            editPoint.Insert(sprintf "#r \"%s\"\r\n" referenceName)

let (|IsFrameworkSpecific|_|) (input:string) =
    let libIndex = input.IndexOf("\\lib\\") + 5
    let modifiedInput = input.Substring(libIndex, input.Length - libIndex)
    match modifiedInput.Contains("\\") with
    | true -> Some ()
    | _ -> None

let UpdateReference path =    
    Directory.GetFiles(GetRepositoryPath path, "*.dll", SearchOption.AllDirectories) 
    |> Seq.filter(fun file -> file.ToLower().Contains("lib"))
    |> Seq.iter(fun file -> 
                    match file with
                    | IsFrameworkSpecific -> 
                        // TODO: Need to make this better
                        if file.ToUpper().Contains("NET40") then
                            AddReferenceToFsx file
                    | _ -> AddReferenceToFsx file )

let InstallPackage packageName = 
    InstallPackages packageName path
    sprintf "%s has been retrieved successfully. Now updating the references." packageName |> SendToFsiConsole
    UpdateReference path
    sprintf "%s has been added successfully." packageName |> SendToFsiConsole 
 
//Example: InstallPackage "FSPowerPack.Core.Community"