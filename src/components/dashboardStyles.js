// src/components/dashboardStyles.js
export const getDashboardStyles = (mode) => {
  if (mode === 'dark') {
    return {
      header: {
        fontWeight: 'bold',
        color: '#fff', // weiße Schrift im Dark Mode
      },
      excelUpload: {
        mt: 3,
        p: 2,
        backgroundColor: '#333', // dunkler Hintergrund
        borderRadius: 1,
      },
      dashboardInputs: {
        mt: 3,
        p: 2,
        backgroundColor: '#444', // etwas hellerer dunkler Hintergrund
        borderRadius: 1,
      },
      templateUpload: {
        mt: 3,
        p: 2,
        backgroundColor: '#555', // noch heller
        borderRadius: 1,
      },
      documentGeneration: {
        mt: 3,
        p: 2,
        backgroundColor: '#666', // dunkler Orangeton-ähnlich
        borderRadius: 1,
        textAlign: 'center',
      },
    };
  } else {
    return {
      header: {
        fontWeight: 'bold',
      },
      excelUpload: {
        mt: 3,
        p: 2,
        backgroundColor: '#e3f2fd',
        borderRadius: 1,
      },
      dashboardInputs: {
        mt: 3,
        p: 2,
        backgroundColor: '#f1f8e9',
        borderRadius: 1,
      },
      templateUpload: {
        mt: 3,
        p: 2,
        backgroundColor: '#ede7f6',
        borderRadius: 1,
      },
      documentGeneration: {
        mt: 3,
        p: 2,
        backgroundColor: '#fff3e0',
        borderRadius: 1,
        textAlign: 'center',
      },
    };
  }
};
