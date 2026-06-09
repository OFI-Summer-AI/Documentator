import '@ofi-summerai/markui/style.css'

import { Button } from '@ofi-summerai/markui/primitives'
import { KPICard } from '@ofi-summerai/markui/components'
import { TrendingUp } from 'lucide-react'

export default function Dashboard() {
  return (
    <div className="p-6 space-y-4">
      <div className="grid grid-cols-3 gap-4">
        <KPICard
          title="Total Reach"
          value="124,500"
          trend={12.4}
          trendLabel="vs last month"
          icon={<TrendingUp size={18} />}
        />
        <KPICard
          title="Engagement Rate"
          value="4.8%"
          trend={-1.2}
          trendLabel="vs last week"
        />
        <KPICard
          title="Posts Published"
          value={38}
          trend={5.0}
          trendLabel="this month"
        />
      </div>

      <div className="flex gap-2">
        <Button>Create Post</Button>
        <Button variant="outline">View Analytics</Button>
        <Button variant="ghost">Settings</Button>
      </div>
    </div>
  )
}