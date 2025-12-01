// Import default arrays from each module
import userTools from "./user.js";
import driveTools from "./drive.js";
import mailTools from "./mail.js";

// Export aggregated array as default
const allTools = [...userTools, ...driveTools, ...mailTools];

export default allTools;
