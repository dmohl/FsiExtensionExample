open Microsoft.FSharp.Control  
  
let rec job = async {     
        for i in 1 .. 20 do    
            printfn "doing some work"  
            do! Async.Sleep 300  
            worker.ReportProgress i  
    }  
    and worker : AsyncWorker<_> = AsyncWorker(job)  
  
worker.ProgressChanged.Add(fun jobNumber -> printfn "job %d completed" jobNumber)   
worker.Error.Add(fun err -> printfn "Error: %A" err.Message)  
worker.Completed.Add(fun _ -> printfn "All jobs have completed")  
worker.Canceled.Add(fun _ -> printfn "Jobs have been canceled")  
  
worker.RunAsync() |> ignore 