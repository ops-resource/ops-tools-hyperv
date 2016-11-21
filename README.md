# Ops-Tools-HyperV

The ops-tools-hyperv repository contains all the scripts and tools to interact with the Hyper-V hypervisor.


## Installation instructions

The different scripts and tools can be obtained from [NuGet.org](https://nuget.org).


## Contributing

To build the project invoke MsBuild on the `entrypoint.msbuild` script in the repository root directory with the `build` target. This will package the scripts and create
the NuGet packages and ZIP archives. Final artifacts will be placed in the `build\deploy` directory.

The build script assumes that:

* The connection to the repository is available so that the version number can be obtained via [GitVersion](https://github.com/ParticularLabs/GitVersion).
* The NuGet command line executable is available from the PATH
