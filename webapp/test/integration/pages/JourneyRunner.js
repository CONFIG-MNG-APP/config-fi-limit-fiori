sap.ui.define([
    "sap/fe/test/JourneyRunner",
	"zgsp26/conf/mng/filimit/confmngfefilimit/test/integration/pages/FILimitConfMain"
], function (JourneyRunner, FILimitConfMain) {
    'use strict';

    var runner = new JourneyRunner({
        launchUrl: sap.ui.require.toUrl('zgsp26/conf/mng/filimit/confmngfefilimit') + '/test/flp.html#app-preview',
        pages: {
			onTheFILimitConfMain: FILimitConfMain
        },
        async: true
    });

    return runner;
});
