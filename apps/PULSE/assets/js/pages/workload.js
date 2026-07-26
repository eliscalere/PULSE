function buildWorkloadByAssignee(db) {
  const workloadByAssignee = {};
  (db.members || []).forEach((member) => {
    workloadByAssignee[member.name] = { member, effort: 0, taskCount: 0, tasks: [] };
  });

  const unassigned = { member: { name: "Unassigned", role: "" }, effort: 0, taskCount: 0, tasks: [] };

  Object.keys(db.ganttTasks || {}).forEach((projId) => {
    const tasks = db.ganttTasks[projId] || [];
    const projName = (db.projects.find((p) => p.id === projId) || {}).name || projId;
    tasks.forEach((task) => {
      if (task.status === "Done" || task.health === "Done") return;
      const effort = Number(task.estimatedEffort) || 0;
      const assigneeName = task.assignee || "Unassigned";

      let assigneeData = unassigned;
      if (assigneeName !== "Unassigned") {
        if (!workloadByAssignee[assigneeName]) {
          workloadByAssignee[assigneeName] = { member: { name: assigneeName, role: "" }, effort: 0, taskCount: 0, tasks: [] };
        }
        assigneeData = workloadByAssignee[assigneeName];
      }

      assigneeData.effort += effort;
      assigneeData.taskCount += 1;
      assigneeData.tasks.push({ ...task, projName });
    });
  });

  const allAssignees = Object.values(workloadByAssignee).filter((d) => d.taskCount > 0).concat(unassigned.taskCount > 0 ? [unassigned] : []);
  allAssignees.sort((a, b) => b.effort - a.effort);
  return allAssignees;
}

function renderWorkloadTableHtml(assignees) {
  return `
    <div class="aewttr-card">
      <table class="aewttr-table">
        <thead>
          <tr>
            <th>Assignee</th>
            <th>Role</th>
            <th>Active Tasks</th>
            <th>Est. Effort (hrs)</th>
          </tr>
        </thead>
        <tbody>
          ${assignees.length ? assignees.map((d) => {
            const member = (window.AEWTTR.db.members || []).find((m) => m.name === d.member.name);
            return `
            <tr ${member ? `data-person-id="${escapeHtml(member.id)}" style="cursor:pointer;"` : ""}>
              <td>
                <div class="people-row-person">
                  ${userAvatarHtml({ name: d.member.name, email: member && member.email, size: 28 })}
                  <strong>${escapeHtml(d.member.name)}</strong>
                </div>
              </td>
              <td>${escapeHtml(d.member.role || "—")}</td>
              <td>${d.taskCount}</td>
              <td>${d.effort > 0 ? `<strong>${d.effort}</strong> hrs` : "—"}</td>
            </tr>`;
          }).join("") : `<tr><td colspan="4"><div class="empty-state" style="padding:40px;">No active tasks found across projects.</div></td></tr>`}
        </tbody>
      </table>
    </div>
  `;
}

PAGE_RENDERERS.workload = function () {
  window.AEWTTR.state.overviewView = "Workload";
  navigate("overview");
};
