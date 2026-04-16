using System;
using System.Windows;
using Velopack;

namespace PDFProcessor
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            
            // Velopack 업데이트 적용
            VelopackApp.Build()
                .OnFirstRun((v) => {
                    // 첫 실행 시 동작
                    MessageBox.Show($"PDF 통합 처리기가 설치되었습니다!", 
                        "설치 완료", MessageBoxButton.OK, MessageBoxImage.Information);
                })
                .Run();
        }
    }
}
