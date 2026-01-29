import * as React from "react"
import { format, subDays, subMonths, subYears, startOfDay, endOfDay } from "date-fns"
import { FiCalendar } from "react-icons/fi"
import { cn } from "@/lib/utils"
import { Button } from "@/components/ui/button"
import { Calendar } from "@/components/ui/calendar"
import {
    Popover,
    PopoverContent,
    PopoverTrigger,
} from "@/components/ui/popover"

const DATE_PRESETS = [
    { label: "Custom", getValue: () => null }, // Special case for manual selection
    { label: "Today", getValue: () => ({ from: startOfDay(new Date()), to: endOfDay(new Date()) }) },
    { label: "Yesterday", getValue: () => ({ from: startOfDay(subDays(new Date(), 1)), to: endOfDay(subDays(new Date(), 1)) }) },
    { label: "Last 7 days", getValue: () => ({ from: startOfDay(subDays(new Date(), 7)), to: endOfDay(new Date()) }) },
    { label: "Last 30 days", getValue: () => ({ from: startOfDay(subDays(new Date(), 30)), to: endOfDay(new Date()) }) },
    { label: "Last 6 months", getValue: () => ({ from: startOfDay(subMonths(new Date(), 6)), to: endOfDay(new Date()) }) },
    { label: "Last year", getValue: () => ({ from: startOfDay(subYears(new Date(), 1)), to: endOfDay(new Date()) }) },
    { label: "All Time", getValue: () => ({ from: undefined, to: undefined }) },
]

export function DateRangePicker({ value, onChange, placeholder = "Select date range", className }) {
    const [isOpen, setIsOpen] = React.useState(false)
    const [range, setRange] = React.useState(value || { from: undefined, to: undefined })
    const [selectedPreset, setSelectedPreset] = React.useState(null)
    const [leftMonth, setLeftMonth] = React.useState(new Date())
    const [rightMonth, setRightMonth] = React.useState(() => {
        const next = new Date()
        next.setMonth(next.getMonth() + 1)
        return next
    })

    React.useEffect(() => {
        if (value) {
            setRange(value)
        }
    }, [value])

    const handlePresetClick = (preset, index) => {
        // Skip if Custom is clicked (it's only for display)
        if (index === 0) return

        const newRange = preset.getValue()
        setRange(newRange)
        setSelectedPreset(index)
    }

    const handleApply = () => {
        if (onChange) {
            onChange(range)
        }
        setIsOpen(false)
    }

    const handleCancel = () => {
        setRange(value || { from: undefined, to: undefined })
        setSelectedPreset(null)
        setIsOpen(false)
    }

    const formatDateRange = () => {
        if (!range?.from) return placeholder
        if (!range.to) return format(range.from, "MMM d, yyyy")

        // Check if both dates are in the same year
        const sameYear = range.from.getFullYear() === range.to.getFullYear()

        if (sameYear) {
            return `${format(range.from, "MMM d")} - ${format(range.to, "MMM d, yyyy")}`
        } else {
            // Different years: show year for both dates
            return `${format(range.from, "MMM d, yyyy")} - ${format(range.to, "MMM d, yyyy")}`
        }
    }

    return (
        <Popover open={isOpen} onOpenChange={setIsOpen}>
            <PopoverTrigger asChild>
                <Button
                    variant="outline"
                    className={cn(
                        "justify-start text-left font-normal px-4 py-2 h-auto border-gray-300",
                        !range?.from && "text-gray-400",
                        className
                    )}
                >
                    <FiCalendar className="mr-2 h-4 w-4" />
                    {formatDateRange()}
                </Button>
            </PopoverTrigger>
            <PopoverContent className="w-auto p-0 max-w-fit" align="start">
                <div className="flex">
                    {/* Left Sidebar - Presets */}
                    <div className="border-r border-gray-200 p-3 space-y-1 w-36 shrink-0">
                        {DATE_PRESETS.map((preset, index) => (
                            <button
                                key={preset.label}
                                onClick={() => handlePresetClick(preset, index)}
                                className={cn(
                                    "w-full text-left px-3 py-2 text-sm rounded-md transition-colors whitespace-nowrap",
                                    selectedPreset === index
                                        ? "bg-blue-50 text-blue-700 font-medium"
                                        : "text-gray-700 hover:bg-gray-100"
                                )}
                            >
                                {preset.label}
                            </button>
                        ))}
                    </div>

                    {/* Right Side - Dual Calendar */}
                    <div className="p-4">
                        <div className="flex gap-6">
                            {/* Left Calendar */}
                            <Calendar
                                mode="range"
                                selected={range}
                                onSelect={(newRange) => {
                                    setRange(newRange)
                                    setSelectedPreset(0) // Set to Custom when manually selecting
                                }}
                                month={leftMonth}
                                onMonthChange={setLeftMonth}
                                captionLayout="dropdown"
                                initialFocus
                            />

                            {/* Right Calendar */}
                            <Calendar
                                mode="range"
                                selected={range}
                                onSelect={(newRange) => {
                                    setRange(newRange)
                                    setSelectedPreset(0) // Set to Custom when manually selecting
                                }}
                                month={rightMonth}
                                onMonthChange={setRightMonth}
                                captionLayout="dropdown"
                            />
                        </div>

                        {/* Action Buttons */}
                        <div className="flex justify-end gap-2 mt-4 pt-3 border-t border-gray-200">
                            <Button
                                variant="outline"
                                size="sm"
                                onClick={handleCancel}
                                className="text-gray-700 h-8 px-3"
                            >
                                Cancel
                            </Button>
                            <Button
                                size="sm"
                                onClick={handleApply}
                                className="bg-[#1A72B9] hover:bg-[#1565C0] h-8 px-4"
                            >
                                Apply
                            </Button>
                        </div>
                    </div>
                </div>
            </PopoverContent>
        </Popover>
    )
}
