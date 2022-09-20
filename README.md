# CST-simulation-macros
Repository to store various macros useful for CST simulations

If you would like to develop your own macro to script task parameters in CST, I would recommend using [this link](https://space.mit.edu/RADIO/CST_online/mergedProjects/VBA_DES/special_vbacommands/simulationtaskobject.htm) to start. 

# Instructions
**Disclaimer**: These macros are meant to be used together. If you only want to use one in particular, make sure the nomenclature of your tasks respects the norms used in these macros. 

### Norms:
1. All coil tasks are in a sequence called CoilCombinations
2. All coil tasks are named Coil_#, where # is the coil number
3. All excitation tasks are regrouped in a sequence called SAR_Excitations
4. All excitations tasks are named Exc_#1_#2 or Exc_#1_#2_prime, where #1 is the first coil number and #2 the second coil number
5. All extra tasks must be included in the CoilCombinations sequence
6. Extra tasks must contain the word zero if you want your phases to be zero for this task and neg if you want the opposite sign for your phases. If none of these keywords are present, the phases will be set to their normal value.
7. The block name of your schematic must be called MWSSCHEM1 (this is the default block name)

### How to add a macro:
1. Download the zip file from this repository
2. Extract the file where you want to regroup your macros.
3. Change the parameters using Notepad or any other text editor
4. Open CST
5. Click on `Home --> Macros --> Import VBA Macro`
6. Go to the folder where you extracted the macros and select the one you want. <br>
**Note**: The macro will now be uploaded in the sub folder `Model/3d/`. From now on, if you modify the macros in CST's macro editor, they will only be modified in `Model/3d/`. Therefore, you need to reupload the macro if you modified it in your macro folder or downloaded a new version from GitHub. 
7. Run the macro by clicking on `Home --> Macros --> Name of the macro`

### Contributing to this repository
  * If you want to participate in this repository by providing ideas for new macros, improvements for current macros or clarifications for this README, please create an issue or a pull request with a detailed title and description. 
