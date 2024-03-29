# Contributing to BESTEST-GSR

To contribute additions/changes:

* Fork this repo
* Create a new branch `git checkout -b my-changes`
* Commit your changes
* Create a pull request
* Agree to the BESTEST-GSR Contribution Policy below

# BESTEST-GSR Contribution Policy
*Version 1.0*


The Building Energy Simulation Test - Generation Simulation and Reporting (BESTEST-GSR) team welcomes your contribution to the project. You can contribute to BESTEST-GSR  
project in several ways: by using the software, reporting issues, contributing documentation, or 
contributing code back to the project. The GitHub [Contributing to Open Source](https://opensource.guide/how-to-contribute/)
guide provides a good overview. If you contribute code, you agree that your contribution may be 
incorporated into BESTEST-GSR and made available under the BESTEST-GSR license.

The contribution process for BESTEST-GSR is composed of three steps:

1.	**Send consent email**

    In order for us to distribute your code as part of BESTEST-GSR under the BESTEST-GSR 
[license](../LICENSE.md), we’ll need 
your consent. An email acknowledging understanding of these terms and agreeing to them is
all that will be asked of any contributor. Send an email to the BESTEST-GSR project manager (see 
below for the address) including the following text and a list of co-contributors (if any):
        
    *I agree to contribute to BESTEST-GSR. I agree to the following terms and conditions for my 
contributions: First, I agree that I am licensing the copyright to my contributions under 
the terms of the current BESTEST-GSR license. Second, I hereby grant to Alliance for Sustainable 
Energy, LLC, to any successor manager and distributor of BESTEST-GSR appointed by the U.S. 
Department of Energy, and to all recipients of a version of BESTEST-GSR that includes my 
contributions, a non-exclusive, worldwide, royalty-free, irrevocable patent license under 
any patent claims owned by me, or owned by my employer and known to me, that are or will be,
necessarily infringed upon by my contributions alone, or by combination of my contributions 
with the version of BESTEST-GSR to which they are contributed, to make, have made, use, offer to 
sell, sell, import, and otherwise transfer any version of BESTEST-GSR that includes my 
contributions, in source code and object code form. Third, I represent and warrant that I 
am authorized to make the contributions and grant the foregoing license(s). Additionally, 
if, to my knowledge, my employer has rights to intellectual property that covers my 
contributions, I represent and warrant that I have received permission to make these 
contributions and grant the foregoing license(s) on behalf of my employer.*
        
    Once we have your consent on file, you’ll only need to redo it if conditions change (e.g. a 
change of employer).


2.	**Scope agreement and timeline commitment**

    If your contribution is small (e.g. a bug fix), simply submit your contribution via GitHub. 
If you find a bug, first make sure it is not an already known issue, then report it in the GitHub 
[issue tracker](../../../issues) for this repository. If your 
contribution is larger (e.g. a new feature or new functionality/capability), we’ll need to evaluate 
your proposed contribution first. To do that, we need a written description of why you wish to 
contribute to BESTEST-GSR, a detailed description of the project that you are proposing, the 
precise functionalities that you plan to implement as part of the project, and a timeframe for 
implementation (see [here](BESTEST-GSR_Contribution_Proposal_v1.0_2018-04-04.doc) for the template contribution proposal document). After 
we review your materials, we will schedule a meeting or conference call to discuss your 
information in more detail. We may ask you to revise your materials and make changes to it, 
which we will re-review. Before you do any work we must reach prior agreement and written 
approval on project areas, scope, timeframe, expected contents, and functionalities to be 
addressed. 

3.  **Technical contribution process**

    We want BESTEST-GSR to adhere to our established quality standards. As such, we ask that you follow 
the information below. Smaller, non-code contributions may not require as much review as code contributions, 
but all contributions will be reviewed. Code contributions will initially be in a source 
control branch, and then will be merged into the official BESTEST-GSR repository after review and 
approval. Any bugs, either discovered by you, us, or any users will be tracked in our issue 
tracker. We request you that you take full responsibility for correcting bugs. Be aware 
that, unless notified otherwise, the correction of bugs takes precedence over the 
submission or creation of new code.
        
**Timeline** - BESTEST-GSR is currently released publicly four times a year, shortly after major EnergyPlus and OpenStudio releases that it tests. 

**Code Reviews** - You will be working and testing your code in a source control branch. When a 
piece of functionality is complete, tested and working, let us know and we will review your code. 
If the functionality that you contributed is complex, we will ask you for a written design document 
as well. We want your code to follow coding standards, be clear, readable and maintainable, and of 
course it should do what it is supposed to do. We will look for errors, style issues, comments (or 
lack thereof), and any other issues in your code. We will inform you of our comments and we expect 
you to make the recommended changes. New re-reviews may be expected until the code complies with 
our required processes.

**Unit Tests** - We ask that you supply unit tests along with the code that you have written, if applicable. A 
unit test is a program that exercises your code in isolation to verify that it does what it is 
supposed to do. Your unit tests are very important to us. First, they give an indication that your 
code works according to its intended functionality. Second, we execute your unit tests 
automatically along with our unit tests to verify that the overall BESTEST-GSR code continues to work.

**Code Coverage** - We require that your unit tests provide an adequate coverage of the source code 
you are submitting. You will need to design your unit tests in a way that all critical parts of 
the code (at least) are tested and verified.

**Documentation** - Proper documentation is crucial for our users, without it users will not know 
how to use your contribution. We require that you create user documentation so that end users know 
how to use your new functionality.

For further questions or information:

&nbsp;&nbsp;&nbsp;&nbsp;David Goldwasser<br/>
&nbsp;&nbsp;&nbsp;&nbsp;BESTEST-GSR Project Management<br/>
&nbsp;&nbsp;&nbsp;&nbsp;david.goldwasser@nrel.gov<br/>
&nbsp;&nbsp;&nbsp;&nbsp;303.275.3000<br/>
    
BESTEST-GSR is funded by the U.S. Department of Energy’s (DOE) Building Technologies Office (BTO), and 
managed by the National Renewable Energy Laboratory (NREL).

BESTEST-GSR is developed in collaboration with NREL and private firms.

**Documents**
 
[BESTEST-GSR_Contribution_Proposal_v1.0_2022-05-27.doc](BESTEST-GSR_Contribution_Proposal_v1.0_2022-05-27.doc)
