import socket
import uno
import os
import sys
import time

# Initialize the UNO runtime
localContext = uno.getComponentContext()
resolver = localContext.ServiceManager.createInstanceWithContext(
                "com.sun.star.bridge.UnoUrlResolver", localContext)
ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
smgr = ctx.ServiceManager

def insert_macro(model, macro_path):
    # Query the interface
    script_provider_supplier = model.getScriptProviderSupplier()

    if script_provider_supplier is not None:
        # Now you can use the 'script_provider_supplier' object to access the script provider
        # and manage macros
        # For example:
        script_provider = script_provider_supplier.getScriptProvider()
        if script_provider is not None:
            print("successs")
            # Now you can use 'script_provider' to manage macros
            # e.g., script_provider.createScript, script_provider.getModule, etc.
        else:
            print("Script provider not found.")
    else:
        print("XScriptProviderSupplier interface not found.")

    # Access the BasicScript provider
    script_provider_supplier = model.com.sun.star.script.provider.XScriptProviderSupplier
    script_provider = script_provider_supplier.getScriptProvider("vnd.sun.star.script:Standard.Module1?language=Basic&location=document")
    
    if not script_provider:
        raise RuntimeError("Script provider not found.")

    # Access the module container
    module_container = script_provider.getModule("")
    if not module_container:
        raise RuntimeError("Module container not found.")

    # Load the macro from file
    macro_file = open(macro_path, "r")
    macro_content = macro_file.read()
    macro_file.close()

    # Insert the macro into the module
    module_container.insertByName("Module1", macro_content)

def main(template_file, macro_file, output_file):
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    model = desktop.getCurrentComponent()
    # print(dir(model))
    print(dir(model.ScriptContainer.ScriptProvider))

    if model is None:
        print("Error loading the document.")
        return

    insert_macro(model, macro_file)
    time.sleep(25)  # Use sleep instead of wait
    model.storeToURL(os.path.abspath(output_file), ())

if __name__ == "__main__":
    if len(sys.argv) != 4:
        print("Usage: python insert_text.py <template_file> <macro_file> <output_folder>")
        sys.exit(1)

    template_file = sys.argv[1]
    macro_file = sys.argv[2]
    output_file = sys.argv[3]

    main(template_file, macro_file, output_file)
