import xml.etree.ElementTree as ET
import time
import re
from colorama import Fore, Style, init
import os

def applyJmxNamingConventions():
    """
    Script Name: JMXApplyNamingConvention
    Description: this scripts takes as entry a jmx file and apply a naming conventions for TC contrllers,  sampler, extractors, processors
    Modification Date: 2025-04-14
    Version: 1.2
    Created by: Andres Felipe Castano
    Contact: castao_andres@optum.com
    """




    def addSamplers(parent, httpsamplersList):
        for child in parent:
            if child.tag == 'HTTPSamplerProxy':
                httpsamplersList.append(child)
            elif child.tag == 'hashTree':
                addSamplers(child, httpsamplersList)


    def addTCs(parent, dic):
        for i, child in enumerate(parent):
            if child.tag == 'TransactionController':
                dic[child.attrib.get('testname')] = child
            elif child.tag == 'hashTree':
                addTCs(child, dic)


    def addTCs2(parent, dic):
        for i, child in enumerate(parent):
            if child.tag == 'TransactionController':
                dic[child.attrib.get('testname')] = parent[i + 1]
            elif child.tag == 'hashTree':
                addTCs2(child, dic)



    while True:
        file_path = input(Style.RESET_ALL+"Insert the name of the jmx file to process: ")
        if  not os.path.exists(file_path):
            print(Fore.RED + f"\n File not found: {file_path}")
            print(Style.RESET_ALL + "Please make sure you include the extension (e.g., .jmx) and provide the full path if needed.")
            print(" Tip: You can also drag and drop the file into this window and press Enter.")
            print(f" Current directory: {os.getcwd()}")
            continue
        break

    # Load the JMX file
    tree = ET.parse(file_path)
    root = tree.getroot()

    threadGroups = {child.attrib.get('testname'): root[0][1][i + 1] for i, child in enumerate((root[0][1])) if
                    child.tag == 'ThreadGroup'}

    #change TC names by adding the ThreadGroup Name
    for threagroupName in threadGroups:
        TCs = {}
        addTCs(threadGroups[threagroupName], TCs)
        for i, TC in enumerate(TCs):
            TCName = F"{threagroupName}. {str(i).zfill(2)} {TCs[TC].get('testname')}."
            TCs[TC].set('testname', TCName)

    #change the sampler name
    for threagroupName in threadGroups:
        TCs2 = {child.attrib.get('testname'): threadGroups[threagroupName][i + 1] for i, child in
                enumerate(threadGroups[threagroupName]) if child.tag == 'TransactionController'}
        TCs = {}
        addTCs2(threadGroups[threagroupName], TCs)
        for tcName in TCs:
            samplers = []
            addSamplers(TCs[tcName], samplers)
            for i, sampler in enumerate(samplers):
                samplerMethod = [child.text for child in sampler if child.get("name") == 'HTTPSampler.method'][0]
                # Next try is for those cases where the host text box is empty in the sampler.
                # It happens mainly with the full url is taken from previous requests or from default config element
                try:
                    samplerHost = [child.text for child in sampler if child.get("name") == 'HTTPSampler.domain'][0]
                except IndexError:
                    samplerHost = ''
                samplerHost = samplerHost.replace('$', '')
                samplerPath = [child.text for child in sampler if child.get("name") == 'HTTPSampler.path'][0]
                if len(samplerPath) > 50:
                    samplerPath = "/...." + "/".join(samplerPath.split('/')[-3:])
                samplerPath = samplerPath.replace('$', '')
                samplerName = f"{tcName}-{i} {samplerMethod} {samplerHost}{samplerPath}"
                #print(samplerName)
                sampler.set('testname', samplerName)

    # Store original and new varaibles
    rename_dict={}
    # Regex extractor
    regex_extractors = root.findall('.//RegexExtractor')
    for extractor in regex_extractors:
        refname = extractor.find('.//stringProp[@name="RegexExtractor.refname"]')
        default = extractor.find('.//stringProp[@name="RegexExtractor.default"]')
        original_name=refname.text
        if not refname.text.capitalize().startswith('C_'):
            new_name = 'C_' + original_name
            rename_dict[original_name]=new_name
            refname.text = new_name
        default.text = refname.text + "__not_found"
        extractor.get('testname')
        extractorName = f'{extractor.get("testclass")} {refname.text}'
        extractor.set('testname', extractorName)

    #css extractor
    css_extractors = root.findall('.//HtmlExtractor')
    for extractor in css_extractors:
        refname = extractor.find('.//stringProp[@name="HtmlExtractor.refname"]')
        default = extractor.find('.//stringProp[@name="HtmlExtractor.default"]')
        original_name = refname.text
        if not refname.text.capitalize().startswith('C_'):
            new_name = 'C_' + original_name
            rename_dict[original_name] = new_name
            refname.text = new_name
        default.text = refname.text + "__not_found"
        extractor.get('testname')
        extractorName = f'{extractor.get("testclass")} {refname.text}'
        extractor.set('testname', extractorName)

    # json extractor
    json_extractors = root.findall('.//JSONPostProcessor')
    for extractor in json_extractors:
        refname = extractor.find('.//stringProp[@name="JSONPostProcessor.referenceNames"]')
        if extractor.find('.//stringProp[@name="JSONPostProcessor.defaultValues"]') == None:
            newTag = ET.Element('stringProp', name='JSONPostProcessor.defaultValues')
            newTag.text = refname.text + "__not_found"
            extractor.append(newTag)
        else:
            default = extractor.find('.//stringProp[@name="JSONPostProcessor.defaultValues"]')
            default.text = refname.text + "__not_found"
        if not refname.text.capitalize().startswith('C_'):
            original_name = refname.text
            new_name = 'C_' + original_name
            rename_dict[original_name] = new_name
            refname.text = new_name
        default.text = refname.text + "__not_found"
        extractor.get('testname')
        extractorName = f'{extractor.get("testclass")} {refname.text}'
        extractor.set('testname', extractorName)

    boundary_extractors = root.findall('.//BoundaryExtractor')
    for extractor in boundary_extractors:
        refname = extractor.find('.//stringProp[@name="BoundaryExtractor.refname"]')
        default = extractor.find('.//stringProp[@name="BoundaryExtractor.default"]')
        original_name = refname.text
        if not refname.text.capitalize().startswith('C_'):
            new_name = 'C_' + original_name
            rename_dict[original_name] = new_name
            refname.text = new_name
        default.text = refname.text + "__not_found"
        extractor.get('testname')
        extractorName = f'{extractor.get("testclass")} {refname.text}'
        extractor.set('testname', extractorName)

    # json extractor
    jSR223PostProcessors = root.findall('.//JSR223PostProcessor')
    for i, jSR223PostProcessor in enumerate(jSR223PostProcessors):
        name = jSR223PostProcessor.get('testname')
        jSR223PostProcessor.set('testname', f'{name}_{i}')



    # Convert XML tree to string for global replacement
    temp_output=file_path.replace('.jmx','_Modified.jmx')
    tree.write(temp_output,encoding='utf-8',xml_declaration=True)
    with open(temp_output,'r',encoding='utf-8') as file:
        content=file.read()



    # Replace occurrences of old variable names with new ones
    for original, new in rename_dict.items():
        # JMeter variables format: ${varName}
        pattern = r'(\$\{' + re.escape(original) + r'\})'
        replacement = '${' + new + '}'
        content = re.sub(pattern, replacement, content)
        pattern = r'(vars\.get\(["\']' + re.escape(original) + r'["\']\))'
        replacement = 'vars.get("' + new + '")'
        content = re.sub(pattern, replacement, content)

    def deleteWhiteLines(content,outFile):
        processedLines = []
        previousWhileLine = False
        for linea in content:
            if linea.strip() == '':
                if not previousWhileLine:
                    processedLines.append(linea)
                    previousWhileLine = True
            else:
                processedLines.append(linea)
                previousWhileLine = False

        with open(outFile, 'w', encoding='utf-8') as f_out:
            f_out.writelines(processedLines)

    deleteWhiteLines(content,temp_output)
    print(f""""✅ JMX file successfully processed!
Output saved as: {temp_output}
 """ )
    time.sleep(2)


if __name__ == "__main__":
    applyJmxNamingConventions()
