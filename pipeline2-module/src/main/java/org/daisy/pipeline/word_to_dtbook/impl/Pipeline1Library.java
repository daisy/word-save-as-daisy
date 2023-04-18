package org.daisy.pipeline.word_to_dtbook.impl;

import java.util.Locale;
import java.util.Date;
import java.text.SimpleDateFormat;
import org.daisy.common.xpath.saxon.ExtensionFunctionProvider;
import org.daisy.common.xpath.saxon.ReflexiveExtensionFunctionProvider;
import org.osgi.service.component.annotations.Component;

//
// The contents of this file are subject to the Mozilla Public License Version 1.0 (the "License");
// you may not use this file except in compliance with the License. You may obtain a copy of the
// License at http://www.mozilla.org/MPL/
//
// Software distributed under the License is distributed on an "AS IS" basis,
// WITHOUT WARRANTY OF ANY KIND, either express or implied.
// See the License for the specific language governing rights and limitations under the License.
//
// The Original Code is: all this file.
//
// The Initial Developer of the Original Code is Michael H. Kay.
//
// Portions created by Norman Walsh are Copyright (C) Mark Logic Corporation. All Rights Reserved.
//
// Contributor(s): Norman Walsh.
//

public class Pipeline1Library {

    @Component(
            name = "Pipeline1Library",
            service = { ExtensionFunctionProvider.class }
    )
    public static class Provider extends ReflexiveExtensionFunctionProvider {
        public Provider() {
            super(Pipeline1Library.class);
        }
    }

    public String getDefaultLocale(){
        return Locale.getDefault().toString().replace('_', '-');
    }

    public String getDate(){
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
        return sdf.format(new Date());
    }

}
