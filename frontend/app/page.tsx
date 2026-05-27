"use client";

import { useState } from "react";

export default function Home() {

  const [fileName, setFileName] = useState("");
  const [selectedFile, setSelectedFile] = useState<File | null>(null);

  const [summary, setSummary] = useState<any>(null);

  const [failedChecks, setFailedChecks] = useState<any[]>([]);
  const [passedChecks, setPassedChecks] = useState<any[]>([]);

  const handleFile = (file: File) => {

    setSelectedFile(file);
    setFileName(file.name);
  };

  const handleFileChange = (
    event: React.ChangeEvent<HTMLInputElement>
  ) => {

    const file = event.target.files?.[0];

    if (file) {
      handleFile(file);
    }
  };

  const handleDrop = (
    event: React.DragEvent<HTMLDivElement>
  ) => {

    event.preventDefault();

    const file = event.dataTransfer.files?.[0];

    if (file) {
      handleFile(file);
    }
  };

  const runQC = async () => {

    if (!selectedFile) return;

    const formData = new FormData();

    formData.append("file", selectedFile);

    const response = await fetch(
      "http://127.0.0.1:8000/upload",
      {
        method: "POST",
        body: formData,
      }
    );

    const data = await response.json();

    setSummary(data.summary);

    const failed = data.results.filter(
      (item: any) => item.status === "FAILED"
    );

    const passed = data.results.filter(
      (item: any) => item.status === "PASSED"
    );

    setFailedChecks(failed);
    setPassedChecks(passed);
  };

  return (

    <main className="min-h-screen bg-gray-100 p-10">

      <div className="max-w-7xl mx-auto bg-white rounded-2xl shadow-lg p-8">

        <h1 className="text-4xl font-bold mb-6">
          Automatic QC Platform
        </h1>

        <div
          onDrop={handleDrop}
          onDragOver={(e) => e.preventDefault()}
          className="border-4 border-dashed border-blue-400 rounded-xl p-12 text-center bg-blue-50"
        >

          <p className="text-xl mb-4">
            Drag & Drop Excel File Here
          </p>

          <p className="mb-6 text-gray-600">
            or
          </p>

          <label className="inline-block cursor-pointer bg-blue-600 text-white px-6 py-3 rounded-xl hover:bg-blue-700 transition">

            Choose Excel File

            <input
              type="file"
              accept=".xlsx"
              onChange={handleFileChange}
              className="hidden"
            />

          </label>

          <br />
          <br />

          <button
            onClick={runQC}
            className="bg-black text-white px-6 py-3 rounded-xl hover:bg-gray-800 transition"
          >
            Run QC
          </button>

          {fileName && (
            <p className="mt-6 text-green-600 font-semibold">
              Selected file: {fileName}
            </p>
          )}

        </div>

        {summary && (

          <div className="grid grid-cols-3 gap-6 mt-10">

            <div className="bg-blue-100 p-6 rounded-xl border-2 border-blue-500">
              <h2 className="text-xl font-bold">
                Total Checks
              </h2>

              <p className="text-3xl mt-2">
                {summary.total}
              </p>
            </div>

            <div className="bg-green-100 p-6 rounded-xl border-2 border-green-500">
              <h2 className="text-xl font-bold">
                Passed
              </h2>

              <p className="text-3xl mt-2">
                {summary.passed}
              </p>
            </div>

            <div className="bg-red-100 p-6 rounded-xl border-2 border-red-500">
              <h2 className="text-xl font-bold">
                Failed
              </h2>

              <p className="text-3xl mt-2">
                {summary.failed}
              </p>
            </div>

          </div>

        )}

        {summary && (

          <div className="mt-8">

            <a
              href="http://127.0.0.1:8000/download-report"
              target="_blank"
              className="bg-green-600 text-white px-6 py-3 rounded-xl hover:bg-green-700"
            >
              Download QC Report
            </a>

          </div>

        )}

        {failedChecks.length > 0 && (

          <div className="mt-12 text-left">

            <h2 className="text-3xl font-bold text-red-600 mb-6">
              Errors
            </h2>

            <div className="space-y-3 max-h-[500px] overflow-y-auto">

              {failedChecks.map((item, index) => (

                <div
                  key={index}
                  className="bg-red-100 border border-red-400 p-4 rounded-xl"
                >

                  <div className="font-bold">
                    {item.sheet} — Row {item.row}
                  </div>

                  <div>
                    {item.type}
                  </div>

                  <div>
                    {item.message}
                  </div>

                </div>

              ))}

            </div>

          </div>

        )}

        {passedChecks.length > 0 && (

          <div className="mt-12 text-left">

            <h2 className="text-3xl font-bold text-green-600 mb-6">
              Passed
            </h2>

            <div className="space-y-3 max-h-[500px] overflow-y-auto">

              {passedChecks.map((item, index) => (

                <div
                  key={index}
                  className="bg-green-100 border border-green-400 p-4 rounded-xl"
                >

                  <div className="font-bold">
                    {item.sheet} — Row {item.row}
                  </div>

                  <div>
                    {item.type}
                  </div>

                  <div>
                    {item.message}
                  </div>

                </div>

              ))}

            </div>

          </div>

        )}

      </div>

    </main>
  );
}