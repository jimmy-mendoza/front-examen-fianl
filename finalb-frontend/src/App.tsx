import { Navigate, Route, Routes } from "react-router-dom";
import Reportes from "./pages/Reportes";

function App() {
    return (
        <Routes>
            <Route path="/reportes" element={<Reportes />} />
            <Route path="*" element={<Navigate to="/reportes" replace />} />
        </Routes>
    );
}

export default App;
