# WLBGuard
Outlook add-in to guard your and your colleagues work-life balance


Given Outlook resiliency to disable add-ins with load time longer than 1000m, the first .net based add-in might get punished for using .net and loading the clr. Can be Either:

1. Easily fixed: options -> Add-ins -> COM -> Go and see if you have the checkbox for it
2. Workaround with registry:
  https://docs.microsoft.com/en-us/office/vba/outlook/concepts/getting-started/support-for-keeping-add-ins-enabled
  https://blogs.msdn.microsoft.com/emeamsgdev/2017/08/02/outlooks-slow-add-ins-resiliency-logic-and-how-to-always-enable-slow-add-ins/
