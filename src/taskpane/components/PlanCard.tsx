import React from "react";
import { ProposedPlan } from "../App";

interface PlanCardProps {
  plan: ProposedPlan;
  onApprove: () => void;
  onCancel: () => void;
}

export function PlanCard({ plan, onApprove, onCancel }: PlanCardProps) {
  return (
    <div className="plan-card">
      <div className="plan-card-header">
        <span className="plan-card-icon">📋</span>
        <span className="plan-card-title">계획 확인</span>
      </div>
      <div className="plan-card-summary">{plan.summary}</div>
      <ol className="plan-steps-list">
        {plan.steps.map((step) => (
          <li key={step.step_number} className="plan-step-item">
            <span className="plan-step-number">{step.step_number}</span>
            <div className="plan-step-body">
              <span className="plan-step-action">{step.action}</span>
              {step.detail && (
                <span className="plan-step-detail">{step.detail}</span>
              )}
            </div>
          </li>
        ))}
      </ol>
      <div className="plan-card-actions">
        <button className="plan-btn plan-btn-cancel" onClick={onCancel}>
          취소
        </button>
        <button className="plan-btn plan-btn-approve" onClick={onApprove}>
          승인 후 실행
        </button>
      </div>
    </div>
  );
}
