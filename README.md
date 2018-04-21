# Newsstand

A list of sender addresses of newsletters and social updates. The main repository is updated since 10 May 2017 by the owner.

The format:

```
%% comment

@@ Friendly Name
address1@example.com
address2@example.com

++ include.txt

```

The `Friendly Name` applies to both addresses.

The `++ include.txt` syntax makes your (my) life easier. I can now easily manage my private newsletter addresses (e.g., internal addresses used by my employers) as well as public newsletter addresses.

## Update-MarkAsToBeSweptRule.ps1

This PowerShell script is a simple application using [Microsoft Graph](https://developer.microsoft.com/en-us/graph) to automatically set rules for you. When I was writing the script, I found `Invoke-RestMethod` a very good friend! For information on how to use the script, `Get-Help` the PowerShell way.

## MIT License

Copyright © 2018 by Gee Law

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

- The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

> THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
