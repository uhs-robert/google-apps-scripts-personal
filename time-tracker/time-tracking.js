// time-tracker/time-tracking.js

// Processes time calculation based on Status
const calculateTime = ({
  newStatus,
  prevStatus,
  cellStartTime,
  cellElapsedTime,
}) => {
  const now = new Date();
  if (newStatus === "In Progress") {
    return cellStartTime.setValue(now);
  } else if (
    ["Pending Entry", "On Hold", "Done", "Not Started", "Skipped"].includes(
      newStatus,
    ) &&
    prevStatus === "In Progress"
  ) {
    return setElapsedTime({
      startTime: new Date(cellStartTime.getValue()),
      endTime: now,
      cellElapsedTime,
    });
  }
};

// Calculates elapsed time between two dates
const setElapsedTime = ({ startTime, endTime, cellElapsedTime }) => {
  const prevElapsedTime = cellElapsedTime.getValue();
  const newElapsedTime = (endTime - startTime) / (1000 * 60); // Convert to minutes
  cellElapsedTime.setValue(prevElapsedTime + newElapsedTime);
};
