import { useRef, useState, useEffect, useCallback } from 'react'
import { motion, useInView } from 'motion/react'
import './AnimatedList.css'

const AnimatedItem = ({ children, delay = 0, index, onMouseEnter, onClick }) => {
  const ref = useRef(null)
  const inView = useInView(ref, { amount: 0.5, once: false })
  return (
    <motion.div
      ref={ref}
      data-index={index}
      onMouseEnter={onMouseEnter}
      onClick={onClick}
      initial={{ scale: 0.9, opacity: 0 }}
      animate={inView ? { scale: 1, opacity: 1 } : { scale: 0.9, opacity: 0 }}
      transition={{ duration: 0.2, delay }}
      style={{ marginBottom: '0.5rem', cursor: onClick ? 'pointer' : 'default' }}
    >
      {children}
    </motion.div>
  )
}

export default function AnimatedList({
  items = [],
  renderItem,
  onItemSelect,
  showGradients = true,
  enableArrowNavigation = true,
  className = '',
  itemClassName = '',
  displayScrollbar = true,
  initialSelectedIndex = -1,
  maxHeight = 420,
}) {
  const listRef = useRef(null)
  const [selectedIndex, setSelectedIndex] = useState(initialSelectedIndex)
  const [keyboardNav, setKeyboardNav] = useState(false)
  const [topGradientOpacity, setTopGradientOpacity] = useState(0)
  const [bottomGradientOpacity, setBottomGradientOpacity] = useState(1)

  const handleItemMouseEnter = useCallback((index) => setSelectedIndex(index), [])
  const handleItemClick = useCallback(
    (item, index) => {
      setSelectedIndex(index)
      onItemSelect?.(item, index)
    },
    [onItemSelect]
  )

  const handleScroll = useCallback((e) => {
    const { scrollTop, scrollHeight, clientHeight } = e.target
    setTopGradientOpacity(Math.min(scrollTop / 50, 1))
    const bottomDistance = scrollHeight - (scrollTop + clientHeight)
    setBottomGradientOpacity(scrollHeight <= clientHeight ? 0 : Math.min(bottomDistance / 50, 1))
  }, [])

  useEffect(() => {
    if (!enableArrowNavigation) return
    const handleKeyDown = (e) => {
      if (!listRef.current?.matches(':hover, :focus-within')) return
      if (e.key === 'ArrowDown') {
        e.preventDefault()
        setKeyboardNav(true)
        setSelectedIndex((prev) => Math.min(prev + 1, items.length - 1))
      } else if (e.key === 'ArrowUp') {
        e.preventDefault()
        setKeyboardNav(true)
        setSelectedIndex((prev) => Math.max(prev - 1, 0))
      } else if (e.key === 'Enter' && selectedIndex >= 0) {
        onItemSelect?.(items[selectedIndex], selectedIndex)
      }
    }
    window.addEventListener('keydown', handleKeyDown)
    return () => window.removeEventListener('keydown', handleKeyDown)
  }, [items, selectedIndex, onItemSelect, enableArrowNavigation])

  useEffect(() => {
    if (!keyboardNav || selectedIndex < 0 || !listRef.current) return
    const container = listRef.current
    const selectedItem = container.querySelector(`[data-index="${selectedIndex}"]`)
    if (selectedItem) {
      const extraMargin = 40
      const itemTop = selectedItem.offsetTop
      const itemBottom = itemTop + selectedItem.offsetHeight
      if (itemTop < container.scrollTop + extraMargin) {
        container.scrollTo({ top: itemTop - extraMargin, behavior: 'smooth' })
      } else if (itemBottom > container.scrollTop + container.clientHeight - extraMargin) {
        container.scrollTo({ top: itemBottom - container.clientHeight + extraMargin, behavior: 'smooth' })
      }
    }
    setKeyboardNav(false)
  }, [selectedIndex, keyboardNav])

  return (
    <div className={`animated-list-container ${className}`}>
      <div
        ref={listRef}
        tabIndex={0}
        className={`animated-list ${!displayScrollbar ? 'no-scrollbar' : ''}`}
        style={{ maxHeight }}
        onScroll={handleScroll}
      >
        {items.map((item, index) => (
          <AnimatedItem
            key={index}
            delay={Math.min(index * 0.04, 0.4)}
            index={index}
            onMouseEnter={() => handleItemMouseEnter(index)}
            onClick={onItemSelect ? () => handleItemClick(item, index) : undefined}
          >
            <div className={`animated-list-item ${selectedIndex === index ? 'selected' : ''} ${itemClassName}`}>
              {renderItem ? renderItem(item, index) : <p className="animated-list-item-text">{String(item)}</p>}
            </div>
          </AnimatedItem>
        ))}
      </div>
      {showGradients && (
        <>
          <div className="animated-list-top-gradient" style={{ opacity: topGradientOpacity }} />
          <div className="animated-list-bottom-gradient" style={{ opacity: bottomGradientOpacity }} />
        </>
      )}
    </div>
  )
}
