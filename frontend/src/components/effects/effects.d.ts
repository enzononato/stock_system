declare module '@/components/effects/Galaxy' {
  import { FC } from 'react'
  const Galaxy: FC<Record<string, any>>
  export default Galaxy
}
declare module '@/components/effects/SpecularButton' {
  import { FC, ReactNode } from 'react'
  interface SpecularButtonProps {
    children?: ReactNode
    size?: 'sm' | 'md' | 'lg'
    radius?: number
    tint?: string
    tintOpacity?: number
    blur?: number
    textColor?: string
    lineColor?: string
    baseColor?: string
    intensity?: number
    shineSize?: number
    shineFade?: number
    thickness?: number
    speed?: number
    followMouse?: boolean
    proximity?: number
    autoAnimate?: boolean
    disabled?: boolean
    onClick?: () => void
    className?: string
    type?: 'button' | 'submit' | 'reset'
  }
  const SpecularButton: FC<SpecularButtonProps>
  export default SpecularButton
}
declare module '@/components/effects/BorderGlow' {
  import { FC, ReactNode } from 'react'
  interface BorderGlowProps {
    children?: ReactNode
    className?: string
    edgeSensitivity?: number
    glowColor?: string
    backgroundColor?: string
    borderRadius?: number
    glowRadius?: number
    glowIntensity?: number
    coneSpread?: number
  }
  const BorderGlow: FC<BorderGlowProps>
  export default BorderGlow
}
declare module '@/components/effects/Counter' {
  import { FC, CSSProperties } from 'react'
  interface CounterProps {
    value: number
    fontSize?: number
    padding?: number
    places?: number[]
    gap?: number
    borderRadius?: number
    horizontalPadding?: number
    textColor?: string
    fontWeight?: number | string
    containerStyle?: CSSProperties
    counterStyle?: CSSProperties
    digitStyle?: CSSProperties
    gradientHeight?: number
    gradientFrom?: string
    gradientTo?: string
    topGradientStyle?: CSSProperties
    bottomGradientStyle?: CSSProperties
  }
  const Counter: FC<CounterProps>
  export default Counter
}
declare module '@/components/effects/GlareHover' {
  import { FC, ReactNode, CSSProperties } from 'react'
  interface GlareHoverProps {
    width?: string
    height?: string
    background?: string
    borderRadius?: string
    borderColor?: string
    children?: ReactNode
    glareColor?: string
    glareOpacity?: number
    glareAngle?: number
    glareSize?: number
    transitionDuration?: number
    playOnce?: boolean
    className?: string
    style?: CSSProperties
  }
  const GlareHover: FC<GlareHoverProps>
  export default GlareHover
}
declare module '@/components/effects/AnimatedList' {
  import { FC } from 'react'
  interface AnimatedListProps<T = unknown> {
    items: T[]
    renderItem?: (item: T, index: number) => React.ReactNode
    onItemSelect?: (item: T, index: number) => void
    showGradients?: boolean
    enableArrowNavigation?: boolean
    className?: string
    itemClassName?: string
    displayScrollbar?: boolean
    initialSelectedIndex?: number
    maxHeight?: number
  }
  const AnimatedList: FC<AnimatedListProps<any>>
  export default AnimatedList
}
declare module '@/components/effects/Stepper' {
  import { FC, ReactNode } from 'react'
  interface StepperProps {
    children?: ReactNode
    initialStep?: number
    onStepChange?: (step: number) => void
    onFinalStepCompleted?: () => void
    backButtonText?: string
    nextButtonText?: string
    disableStepIndicators?: boolean
  }
  const Stepper: FC<StepperProps>
  export function Step(props: { children?: ReactNode }): JSX.Element
  export default Stepper
}
