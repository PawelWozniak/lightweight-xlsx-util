# Performance Tests
The performance tests require a dependend package you need to install first in order to store the debug output

These tests should run the tests with a low debug level, so storing the output async is easiest.
Note: These are not complex and detailed tests, it's more to get an idea of how much you can throw at it.

After the package installation deploy the classes in the `performance_tests/classes` folder

## Installation
- Install the *Lightweight - Debug Util* package:  `sf package install --package 04t4K000002WF5uQAG -w 30`
- Run from the project root directory: `sf project deploy start --source-dir performance_tests`
- Execute the `performance_tests/builder/builder_performance_test.apex` for the Builder
- Execute the `performance_tests/parser/parser_performance_test.apex` for the Parser