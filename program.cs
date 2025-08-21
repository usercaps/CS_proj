 <? xml version = "1.0" encoding = "utf-8" ?>
   < Project ToolsVersion = "15.0" xmlns = "http://schemas.microsoft.com/developer/msbuild/2003" >
      
        < Import Project = "$(MSBuildExtensionsPath)\$(MSBuildToolsVersion)\Microsoft.Common.props" Condition = "Exists('$(MSBuildExtensionsPath)\$(MSBuildToolsVersion)\Microsoft.Common.props')" />
         

           < PropertyGroup >
         
             < Configuration Condition = " '$(Configuration)' == '' " > Debug </ Configuration >
          
              < Platform Condition = " '$(Platform)' == '' " > AnyCPU </ Platform >
           
               < ProjectGuid >{ 7E8B9A1C - 6C3E - 4A8F - 9D6A - 13E0B20F4D3A}</ ProjectGuid >
                      
                          < OutputType > WinExe </ OutputType >
                      
                          < RootNamespace > TitleGen </ RootNamespace >
                      
                          < AssemblyName > TitleGen </ AssemblyName >
                      
                          < TargetFrameworkVersion > v4.8 </ TargetFrameworkVersion >
                         
                             < FileAlignment > 512 </ FileAlignment >
                         
                             < AutoGenerateBindingRedirects > true </ AutoGenerateBindingRedirects >
                         
                             < Deterministic > true </ Deterministic >
                         
                           </ PropertyGroup >
                         

                           < PropertyGroup Condition = " '$(Configuration)|$(Platform)' == 'Debug|AnyCPU' " >
                          
                              < PlatformTarget > AnyCPU </ PlatformTarget >
                          
                              < DebugSymbols > true </ DebugSymbols >
                          
                              < DebugType > full </ DebugType >
                          
                              < Optimize > false </ Optimize >
                          
                              < OutputPath > bin\Debug\</ OutputPath >
                             
                                 < DefineConstants > DEBUG; TRACE </ DefineConstants >
                                 
                                   </ PropertyGroup >
                                 

                                   < PropertyGroup Condition = " '$(Configuration)|$(Platform)' == 'Release|AnyCPU' " >
                                  
                                      < PlatformTarget > AnyCPU </ PlatformTarget >
                                  
                                      < DebugType > pdbonly </ DebugType >
                                  
                                      < Optimize > true </ Optimize >
                                  
                                      < OutputPath > bin\Release\</ OutputPath >
                                     
                                         < DefineConstants > TRACE </ DefineConstants >
                                     
                                       </ PropertyGroup >
                                     

                                       < ItemGroup >
                                     
                                         < Reference Include = "System" />
                                      
                                          < Reference Include = "System.Data" />
                                       
                                           < Reference Include = "System.Drawing" />
                                        
                                            < Reference Include = "System.Windows.Forms" />
                                         
                                             < !--Word Interop: путь зависит от версии пакета в packages.config -->
    <!-- ЗАМЕНИ версию/путь ниже на свою (после Restore) -->
    < Reference Include = "Microsoft.Office.Interop.Word" >
 
       < HintPath >$(SolutionDir)packages\Microsoft.Office.Interop.Word.15.0.4795.1000\lib\net20\Microsoft.Office.Interop.Word.dll </ HintPath >
   
         < Private > true </ Private >
   
       </ Reference >
   
     </ ItemGroup >
   

     < ItemGroup >
   
       < Compile Include = "program.cs" />
    
        < !--если у тебя есть другие .cs файлы (например MainForm.cs), добавь их здесь -->
  </ItemGroup>

  <ItemGroup>
    <None Include="packages.config" />
    <!-- при необходимости добавь тут .resx/.ico и т.д. -->
  </ItemGroup>

  <Import Project="$(MSBuildToolsPath)\Microsoft.CSharp.targets" />
</Project>
