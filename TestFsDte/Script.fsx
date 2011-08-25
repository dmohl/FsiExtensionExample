#r "System"
#r "EnvDTE.dll"
#r "EnvDTE80.dll"
#r "Microsoft.VisualStudio.OLE.Interop"
#r "System.Xml.Linq"
#r @"C:\Program Files (x86)\Microsoft ASP.NET\ASP.NET Web Pages\v1.0\Assemblies\NuGet.Core.dll"

open System.IO
open EnvDTE
open EnvDTE80
open System.Runtime.InteropServices
open NuGet

let packageSourceUrl = "https://go.microsoft.com/fwlink/?LinkID=206669"
let packageSource = PackageSource(packageSourceUrl)
let repositoryFactory = PackageRepositoryFactory() 
let repository = repositoryFactory.CreateRepository (packageSource)

let automationObject = Marshal.GetActiveObject("VisualStudio.DTE.10.0")
let dte2 = automationObject :?> DTE2

let documentPath = match dte2.Solution with
            | null -> dte2.ActiveDocument.Path
            | _ -> Path.GetDirectoryName dte2.Solution.FullName

let SendToFsiConsole text = 
    printfn "%s\r\n" text

let GetRepositoryPath path = 
    Path.Combine(path, "packages")

let InstallPackages (repositoryId:string) (destinationPath:string) =
    "Creating the package manager." |> SendToFsiConsole
    let packageManager = PackageManager(repository, DefaultPackagePathResolver documentPath, PhysicalFileSystem(GetRepositoryPath destinationPath))
    "Retrieving the package from the NuGet Repository. This may take several minutes..." |> SendToFsiConsole
    packageManager.InstallPackage repositoryId
    
let AddReferenceToFsx referenceName =
    if dte2.ActiveDocument.Name.Contains ".fsx" then
        let doc = dte2.ActiveDocument.Object "TextDocument" :?> TextDocument
        let editPoint = doc.CreateEditPoint()
        if not ((editPoint.GetText doc.EndPoint).Contains(referenceName)) then
            editPoint.LineDown 1
            editPoint.Insert(sprintf "#r @\"%s\"\r\n" referenceName)

let (|IsFrameworkSpecific|_|) (input:string) =
    let libIndex = input.IndexOf("\\lib\\") + 5
    let modifiedInput = input.Substring(libIndex, input.Length - libIndex)
    match modifiedInput.Contains("\\") with
    | true -> Some ()
    | _ -> None

let UpdateReference path =    
    Directory.GetFiles(GetRepositoryPath path, "*.dll", SearchOption.AllDirectories) 
    |> Seq.filter(fun file -> file.ToLower().Contains("lib") 
                              && (not (file.ToLower().Contains("tools")))
                              && (not (file.ToLower().Contains("content"))) )
    |> Seq.iter(fun file -> 
                    match file with
                    | IsFrameworkSpecific -> 
                        // TODO: Need to make this better
                        if file.ToUpper().Contains("NET40") then
                            AddReferenceToFsx file
                    | _ -> AddReferenceToFsx file )

let InstallPackage packageName = 
    InstallPackages packageName documentPath
    sprintf "%s has been retrieved successfully. Now updating the references." packageName |> SendToFsiConsole
    UpdateReference documentPath
    sprintf "%s has been added successfully." packageName |> SendToFsiConsole 
 
let ShowInstalledPackages () =
    Directory.GetDirectories(GetRepositoryPath documentPath)
    |> Seq.iter (fun dir -> let array = dir.Split([|'\\'|]) |> Array.rev  
                            array.[0] |> SendToFsiConsole )
