import { StepPlaceholder } from "@/components/StepPlaceholder";

export default function AnalyzePage() {
  return (
    <StepPlaceholder
      step="A-2 問題の見える化"
      lead="影響度と頻度で採点し、問題を分類します。既定は2軸（影響度×頻度）で十分です。"
    />
  );
}
