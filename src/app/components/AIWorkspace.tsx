import { useState, type ReactNode, type ChangeEvent } from "react";
import { Textarea } from "@/app/components/ui/textarea";
import { Button } from "@/app/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/app/components/ui/card";
import { Sparkles, Copy, Download, RefreshCw } from "lucide-react";
import { toast } from "sonner";

interface AIWorkspaceProps {
  title: string;
  description: string;
  placeholder: string;
  icon: ReactNode;
}

export function AIWorkspace({ title, description, placeholder, icon }: AIWorkspaceProps) {
  const [input, setInput] = useState("");
  const [output, setOutput] = useState("");
  const [isGenerating, setIsGenerating] = useState(false);

  const handleGenerate = async () => {
    if (!input.trim()) {
      toast.error("Please enter a prompt");
      return;
    }

    setIsGenerating(true);
    setTimeout(() => {
      setOutput(`Generated content based on your input:\n\n${input}\n\n`);
      setIsGenerating(false);
      toast.success("Content generated successfully!");
    }, 2000);
  };

  const handleCopy = () => {
    navigator.clipboard.writeText(output);
    toast.success("Copied to clipboard!");
  };

  const handleDownload = () => {
    const element = document.createElement("a");
    const file = new Blob([output], { type: "text/plain" });
    element.href = URL.createObjectURL(file);
    element.download = "ai-generated-content.txt";
    document.body.appendChild(element);
    element.click();
    document.body.removeChild(element);
    toast.success("Downloaded successfully!");
  };

return (
  <div className="space-y-6">
      <div className="flex items-center gap-3">
        {icon}
        <div>
          <h2 className="text-2xl font-bold">{title}</h2>
          <p className="text-gray-600">{description}</p>
        </div>
      </div>

      <div className="grid md:grid-cols-2 gap-6">
        <Card>
          <CardHeader>
            <CardTitle>Input</CardTitle>
            <CardDescription>Enter your prompt or requirements</CardDescription>
          </CardHeader>
          <CardContent>
            <Textarea
              value={input}
              onChange={(e: ChangeEvent<HTMLTextAreaElement>) => setInput(e.target.value)}
              placeholder={placeholder}
              className="min-h-[300px] resize-none"
            />
            <Button 
              onClick={handleGenerate} 
              className="w-full mt-4"
              disabled={isGenerating}
            >
              {isGenerating ? (
                <>
                  <RefreshCw className="w-4 h-4 mr-2 animate-spin" />
                  Generating...
                </>
              ) : (
                <>
                  <Sparkles className="w-4 h-4 mr-2" />
                  Generate
                </>
              )}
            </Button>
          </CardContent>
        </Card>

        <Card>
          <CardHeader>
            <CardTitle>Output</CardTitle>
            <CardDescription>AI-generated content</CardDescription>
          </CardHeader>
          <CardContent>
            <Textarea
              value={output}
              readOnly
              placeholder="Generated content will appear here..."
              className="min-h-[300px] resize-none bg-gray-50"
            />
            <div className="flex gap-2 mt-4">
              <Button 
                onClick={handleCopy} 
                variant="outline"
                className="flex-1"
                disabled={!output}
              >
                <Copy className="w-4 h-4 mr-2" />
                Copy
              </Button>
              <Button 
                onClick={handleDownload} 
                variant="outline"
                className="flex-1"
                disabled={!output}
              >
                <Download className="w-4 h-4 mr-2" />
                Download
              </Button>
            </div>
          </CardContent>
        </Card>
      </div>
    </div>
  );
}
