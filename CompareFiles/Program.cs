using CompareFiles;

var result =
    FileComparer.Markdown(
        @"C:\users\mark\Downloads\1-introduction.md",
        @"C:\Users\mark\Downloads\learnModules\accelerate-scale-spring-boot-application-azure-cache-redis\includes\1-introduction.md"
    );

foreach (var item in result)
{
    Console.WriteLine(item);
}

return;