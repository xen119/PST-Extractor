import * as React from 'react'
import * as DialogPrimitive from '@radix-ui/react-dialog'
import * as PopoverPrimitive from '@radix-ui/react-popover'
import * as SeparatorPrimitive from '@radix-ui/react-separator'
import * as TooltipPrimitive from '@radix-ui/react-tooltip'
import { X } from 'lucide-react'
import { cn } from '@/lib/utils'

type ButtonVariant = 'default' | 'secondary' | 'ghost' | 'danger'

export interface ButtonProps extends React.ButtonHTMLAttributes<HTMLButtonElement> {
  variant?: ButtonVariant
  iconOnly?: boolean
}

export const Button = React.forwardRef<HTMLButtonElement, ButtonProps>(
  ({ className, variant = 'default', iconOnly = false, ...props }, ref) => {
    const variantClass =
      variant === 'default'
        ? 'button button-primary'
        : variant === 'secondary'
          ? 'button'
          : variant === 'danger'
            ? 'button border-[color:rgba(220,38,38,0.22)] bg-[color:rgba(220,38,38,0.08)] text-[color:var(--danger)] hover:bg-[color:rgba(220,38,38,0.14)]'
            : 'button button-ghost'

    return (
      <button
        ref={ref}
        className={cn(variantClass, iconOnly && 'px-0', className)}
        {...props}
      />
    )
  }
)
Button.displayName = 'Button'

export interface IconButtonProps extends ButtonProps {
  label: string
}

export const IconButton = React.forwardRef<HTMLButtonElement, IconButtonProps>(
  ({ label, className, children, title, ...props }, ref) => {
    return (
      <TooltipPrimitive.Provider delayDuration={120}>
        <TooltipPrimitive.Root>
          <TooltipPrimitive.Trigger asChild>
            <button
              ref={ref}
              className={cn('icon-button', className)}
              aria-label={label}
              title={title || label}
              {...props}
            >
              {children}
            </button>
          </TooltipPrimitive.Trigger>
          <TooltipPrimitive.Portal>
            <TooltipPrimitive.Content
              sideOffset={8}
              className="z-50 rounded-full border border-[color:var(--line)] bg-[color:var(--surface-strong)] px-2.5 py-1 text-xs font-medium text-[color:var(--text)] shadow-soft"
            >
              {label}
              <TooltipPrimitive.Arrow className="fill-[color:var(--surface-strong)]" />
            </TooltipPrimitive.Content>
          </TooltipPrimitive.Portal>
        </TooltipPrimitive.Root>
      </TooltipPrimitive.Provider>
    )
  }
)
IconButton.displayName = 'IconButton'

export const Badge = React.forwardRef<HTMLSpanElement, React.HTMLAttributes<HTMLSpanElement>>(
  ({ className, ...props }, ref) => (
    <span ref={ref} className={cn('badge', className)} {...props} />
  )
)
Badge.displayName = 'Badge'

export const Input = React.forwardRef<HTMLInputElement, React.InputHTMLAttributes<HTMLInputElement>>(
  ({ className, ...props }, ref) => <input ref={ref} className={cn('input', className)} {...props} />
)
Input.displayName = 'Input'

export const Textarea = React.forwardRef<HTMLTextAreaElement, React.TextareaHTMLAttributes<HTMLTextAreaElement>>(
  ({ className, ...props }, ref) => (
    <textarea ref={ref} className={cn('textarea min-h-[96px] resize-y', className)} {...props} />
  )
)
Textarea.displayName = 'Textarea'

export const Separator = React.forwardRef<
  React.ElementRef<typeof SeparatorPrimitive.Root>,
  React.ComponentPropsWithoutRef<typeof SeparatorPrimitive.Root>
>(({ className, ...props }, ref) => (
  <SeparatorPrimitive.Root
    ref={ref}
    className={cn('shrink-0 bg-[color:var(--line)]', className)}
    {...props}
  />
))
Separator.displayName = 'Separator'

export const TooltipProvider = TooltipPrimitive.Provider
export const Tooltip = TooltipPrimitive.Root
export const TooltipTrigger = TooltipPrimitive.Trigger
export const TooltipContent = TooltipPrimitive.Content

export const Dialog = DialogPrimitive.Root
export const DialogTrigger = DialogPrimitive.Trigger
export const DialogPortal = DialogPrimitive.Portal
export const DialogClose = DialogPrimitive.Close
export const DialogOverlay = React.forwardRef<
  React.ElementRef<typeof DialogPrimitive.Overlay>,
  React.ComponentPropsWithoutRef<typeof DialogPrimitive.Overlay>
>(({ className, ...props }, ref) => (
  <DialogPrimitive.Overlay
    ref={ref}
    className={cn('fixed inset-0 z-50 bg-black/40 backdrop-blur-sm', className)}
    {...props}
  />
))
DialogOverlay.displayName = 'DialogOverlay'

export const DialogContent = React.forwardRef<
  React.ElementRef<typeof DialogPrimitive.Content>,
  React.ComponentPropsWithoutRef<typeof DialogPrimitive.Content> & { showCloseButton?: boolean }
>(({ className, children, showCloseButton = true, ...props }, ref) => (
  <DialogPrimitive.Portal>
    <DialogOverlay />
    <DialogPrimitive.Content
      ref={ref}
      className={cn(
        'fixed left-1/2 top-1/2 z-50 max-h-[88vh] w-[min(96vw,1100px)] -translate-x-1/2 -translate-y-1/2 overflow-hidden rounded-3xl border border-[color:var(--line)] bg-[color:var(--surface-strong)] shadow-soft outline-none',
        className
      )}
      {...props}
    >
      {children}
      {showCloseButton ? (
        <DialogPrimitive.Close className="absolute right-4 top-4 rounded-full border border-[color:var(--line)] bg-[color:var(--surface-strong)] p-2 text-[color:var(--text)] transition hover:bg-[color:var(--surface-soft)] focus:outline-none focus:ring-2 focus:ring-[color:var(--focus-ring)]">
          <X className="h-4 w-4" />
          <span className="sr-only">Close</span>
        </DialogPrimitive.Close>
      ) : null}
    </DialogPrimitive.Content>
  </DialogPrimitive.Portal>
))
DialogContent.displayName = 'DialogContent'

export const DialogTitle = React.forwardRef<
  React.ElementRef<typeof DialogPrimitive.Title>,
  React.ComponentPropsWithoutRef<typeof DialogPrimitive.Title>
>(({ className, ...props }, ref) => (
  <DialogPrimitive.Title
    ref={ref}
    className={cn('text-lg font-semibold text-[color:var(--text)]', className)}
    {...props}
  />
))
DialogTitle.displayName = 'DialogTitle'

export const DialogDescription = React.forwardRef<
  React.ElementRef<typeof DialogPrimitive.Description>,
  React.ComponentPropsWithoutRef<typeof DialogPrimitive.Description>
>(({ className, ...props }, ref) => (
  <DialogPrimitive.Description
    ref={ref}
    className={cn('text-sm text-[color:var(--muted)]', className)}
    {...props}
  />
))
DialogDescription.displayName = 'DialogDescription'

export const Popover = PopoverPrimitive.Root
export const PopoverTrigger = PopoverPrimitive.Trigger
export const PopoverClose = PopoverPrimitive.Close
export const PopoverContent = React.forwardRef<
  React.ElementRef<typeof PopoverPrimitive.Content>,
  React.ComponentPropsWithoutRef<typeof PopoverPrimitive.Content>
>(({ className, align = 'start', sideOffset = 10, ...props }, ref) => (
  <PopoverPrimitive.Portal>
    <PopoverPrimitive.Content
      ref={ref}
      align={align}
      sideOffset={sideOffset}
      className={cn(
        'z-50 w-80 rounded-2xl border border-[color:var(--line)] bg-[color:var(--surface-strong)] p-3 shadow-soft outline-none',
        className
      )}
      {...props}
    />
  </PopoverPrimitive.Portal>
))
PopoverContent.displayName = 'PopoverContent'

export const ScrollArea = React.forwardRef<
  React.ElementRef<'div'>,
  React.HTMLAttributes<HTMLDivElement>
>(({ className, children, ...props }, ref) => (
  <div ref={ref} className={cn('overflow-auto', className)} {...props}>
    {children}
  </div>
))
ScrollArea.displayName = 'ScrollArea'
