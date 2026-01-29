import * as React from "react"
import { format } from "date-fns"
import { FiCalendar } from "react-icons/fi"
import { cn } from "@/lib/utils"
import { Button } from "@/components/ui/button"
import { Calendar } from "@/components/ui/calendar"
import {
    Popover,
    PopoverContent,
    PopoverTrigger,
} from "@/components/ui/popover"

export function DatePicker({ value, onChange, placeholder = "Pick a date", className }) {
    const [date, setDate] = React.useState(value ? new Date(value) : undefined)

    React.useEffect(() => {
        if (value) {
            setDate(new Date(value))
        }
    }, [value])

    const handleSelect = (selectedDate) => {
        setDate(selectedDate)
        if (onChange && selectedDate) {
            // Format as YYYY-MM-DD for form compatibility
            const formatted = format(selectedDate, "yyyy-MM-dd")
            onChange(formatted)
        }
    }

    return (
        <Popover>
            <PopoverTrigger asChild>
                <Button
                    variant="outline"
                    className={cn(
                        "w-full justify-start text-left font-normal px-4 py-2.5 h-auto border-gray-300",
                        !date && "text-gray-400",
                        className
                    )}
                >
                    <FiCalendar className="mr-2 h-4 w-4" />
                    {date ? format(date, "PPP") : <span>{placeholder}</span>}
                </Button>
            </PopoverTrigger>
            <PopoverContent className="w-auto p-0" align="start">
                <Calendar
                    mode="single"
                    selected={date}
                    onSelect={handleSelect}
                    initialFocus
                />
            </PopoverContent>
        </Popover>
    )
}
