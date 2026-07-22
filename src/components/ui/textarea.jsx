import * as React from 'react';
import { cn } from '@/lib/utils';

const Textarea = React.forwardRef(({ className, ...props }, ref) => (
  <textarea
    className={cn(
      'flex min-h-[80px] w-full rounded-md border border-slate-200 bg-white px-3 py-2 text-sm shadow-sm placeholder:text-slate-400 focus-visible:outline-none focus-visible:ring-1 focus-visible:ring-indigo-500 disabled:cursor-not-allowed disabled:opacity-50 resize-y',
      className
    )}
    ref={ref}
    {...props}
  />
));
Textarea.displayName = 'Textarea';

export { Textarea };
