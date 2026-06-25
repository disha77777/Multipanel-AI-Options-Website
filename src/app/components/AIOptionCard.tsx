import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/app/components/ui/card";
import { Button } from "@/app/components/ui/button";
import { LucideIcon } from "lucide-react";

interface AIOptionCardProps {
  title: string;
  description: string;
  icon: LucideIcon;
  features: string[];
  onSelect: () => void;
  iconColor: string;
}

export function AIOptionCard({ 
  title, 
  description, 
  icon: Icon, 
  features, 
  onSelect,
  iconColor 
}: AIOptionCardProps) {
  return (
    <Card className="flex flex-col h-full hover:shadow-lg transition-shadow duration-300">
      <CardHeader>
        <div className={`w-12 h-12 rounded-lg ${iconColor} flex items-center justify-center mb-4`}>
          <Icon className="w-6 h-6 text-white" />
        </div>
        <CardTitle>{title}</CardTitle>
        <CardDescription>{description}</CardDescription>
      </CardHeader>
      <CardContent className="flex-1 flex flex-col justify-between">
        <ul className="space-y-2 mb-6">
          {features.map((feature, index) => (
            <li key={index} className="flex items-start">
              <span className="text-green-500 mr-2">✓</span>
              <span className="text-sm text-gray-600">{feature}</span>
            </li>
          ))}
        </ul>
        <Button onClick={onSelect} className="w-full">
          Get Started
        </Button>
      </CardContent>
    </Card>
  );
}
