import { useState } from "react";
import { AIOptionCard } from "@/app/components/AIOptionCard";
import { AIWorkspace } from "@/app/components/AIWorkspace";
import { 
  Sparkles, 
  FileSpreadsheet, 
  Calculator,
  BarChart3,
  ArrowLeft
} from "lucide-react";
import { Button } from "@/app/components/ui/button";
import { Toaster } from "@/app/components/ui/sonner";

type AIOptionType = 
  | "creative" 
  | "excel" 
  | "formulas" 
  | "analytical"
  | null;

function App() {
  const [selectedOption, setSelectedOption] = useState<AIOptionType>(null);

  const aiOptions = [
    {
      id: "creative" as const,
      title: "Creative Writing",
      description: "Generate stories, poems, and creative content",
      icon: Sparkles,
      iconColor: "bg-purple-500",
      features: [
        "Story and plot generation",
        "Poetry and verse creation",
        "Character development",
        "Creative brainstorming"
      ],
      workspaceProps: {
        title: "Creative Writing Assistant",
        description: "Unleash your creativity with AI-powered writing",
        placeholder: "Describe the creative content you want to generate (e.g., 'Write a short story about a time traveler visiting ancient Rome')",
        icon: <Sparkles className="w-8 h-8 text-purple-500" />
      }
    },
    {
      id: "excel" as const,
      title: "Excel Helper",
      description: "Get help with Excel tasks and spreadsheet solutions",
      icon: FileSpreadsheet,
      iconColor: "bg-green-600",
      features: [
        "Data organization tips",
        "Spreadsheet setup help",
        "Pivot table guidance",
        "Chart recommendations"
      ],
      workspaceProps: {
        title: "Excel Helper Assistant",
        description: "Get expert help with Excel and spreadsheets",
        placeholder: "Describe your Excel task (e.g., 'How do I create a pivot table to analyze sales data by region?')",
        icon: <FileSpreadsheet className="w-8 h-8 text-green-600" />
      }
    },
    {
      id: "formulas" as const,
      title: "Formula Generator",
      description: "Create Excel formulas and functions instantly",
      icon: Calculator,
      iconColor: "bg-blue-600",
      features: [
        "Formula generation",
        "Function explanations",
        "Complex calculations",
        "Formula debugging"
      ],
      workspaceProps: {
        title: "Formula Generator",
        description: "Generate and understand Excel formulas",
        placeholder: "Describe the formula you need (e.g., 'Create a formula to calculate the sum of values in column A where column B equals Sales')",
        icon: <Calculator className="w-8 h-8 text-blue-600" />
      }
    },
    {
      id: "analytical" as const,
      title: "Analytical Writing",
      description: "Write data-driven analysis and reports",
      icon: BarChart3,
      iconColor: "bg-orange-600",
      features: [
        "Data analysis reports",
        "Research findings",
        "Business intelligence",
        "Statistical interpretation"
      ],
      workspaceProps: {
        title: "Analytical Writing Assistant",
        description: "Create comprehensive analytical reports",
        placeholder: "Describe your analysis needs (e.g., 'Write an analysis of quarterly sales trends showing 15% growth')",
        icon: <BarChart3 className="w-8 h-8 text-orange-600" />
      }
    }
  ];

  const selectedAIOption = aiOptions.find(opt => opt.id === selectedOption);

  if (selectedOption && selectedAIOption) {
    return (
      <div className="min-h-screen bg-gradient-to-br from-gray-50 to-gray-100">
        <div className="container mx-auto px-4 py-8">
          <Button 
            onClick={() => setSelectedOption(null)} 
            variant="outline"
            className="mb-6"
          >
            <ArrowLeft className="w-4 h-4 mr-2" />
            Back to Options
          </Button>
          <AIWorkspace {...selectedAIOption.workspaceProps} />
        </div>
        <Toaster />
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 via-purple-50 to-pink-50">
      <div className="container mx-auto px-4 py-12">
        {/* Header */}
        <div className="text-center mb-12">
          <div className="flex items-center justify-center gap-3 mb-4">
            <Sparkles className="w-12 h-12 text-purple-600" />
            <h1 className="text-5xl">AI Writing Studio</h1>
          </div>
          <p className="text-xl text-gray-600 max-w-2xl mx-auto">
            Choose your AI writing assistant and create amazing content in seconds
          </p>
        </div>

        {/* AI Options Grid */}
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6 max-w-6xl mx-auto">
          {aiOptions.map((option) => (
            <AIOptionCard
              key={option.id}
              title={option.title}
              description={option.description}
              icon={option.icon}
              iconColor={option.iconColor}
              features={option.features}
              onSelect={() => setSelectedOption(option.id)}
            />
          ))}
        </div>

        {/* Footer */}
        <div className="text-center mt-16 text-gray-500">
          <p>Powered by AI</p>
        </div>
      </div>
      <Toaster />
    </div>
  );
}

export default App;