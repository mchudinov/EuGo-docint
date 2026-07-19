var builder = DistributedApplication.CreateBuilder(args);

builder.AddProject<Projects.DocInt_Api>("docint");

builder.Build().Run();
