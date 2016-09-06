{
    "targets": [
        {
            "target_name": "addon",
            "sources": [
                "src\*.cpp"
            ],
            "include_dirs": [
                "<!(node -e \"require('nan')\")"
            ],

            "configurations": {
                "Debug": {
                    "msvs_settings": {
                        "VCCLCompilerTool": {
                            "RuntimeLibrary": 3,
                            "ExceptionHandling": 1
                        }
                    }
                },
                "Release": {
                    "msvs_settings": {
                        "VCCLCompilerTool": {
                            "RuntimeLibrary": 2,
                            "ExceptionHandling": 1
                        }
                    }
                }
            }
        }
    ],

}
