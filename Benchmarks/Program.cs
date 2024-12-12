using BenchmarkDotNet.Running;
using System.Reflection;

var t = new Benchmarks.Benchmarks()
{
    Records = 100,
};

BenchmarkRunner.Run(Assembly.GetExecutingAssembly());
