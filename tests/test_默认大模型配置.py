"""验证实际GUI加载本机默认模型，保留当前会话开关并处理损坏配置。"""
import json
import pytest
import gui


def _配置(tmp_path, monkeypatch, data):
    monkeypatch.setenv('LOCALAPPDATA',str(tmp_path))
    path=tmp_path/'银行流水核对工具'/'大模型默认配置.json'
    path.parent.mkdir(parents=True)
    path.write_text(json.dumps(data,ensure_ascii=False),encoding='utf-8-sig')
    return path


def test_启动加载指定模型且对话框可关闭当前会话(tmp_path,monkeypatch):
    path=_配置(tmp_path,monkeypatch,{'enabled':True,'mode':'online','protocol':'chat_completions','base_url':'http://example.test:8000/v1','model':'用户指定模型','api_key':'仅用于测试的密钥','timeout_seconds':600,'candidate_limit':5})
    before=path.read_bytes()
    app=gui.ReconciliationApp()
    try:
        app.withdraw()
        assert app.llm_config.enabled
        assert app.llm_config.model=='用户指定模型'
        assert app.llm_config.timeout_seconds==600
        assert '已启用' in app.llm_status_var.get() and '用户指定模型' in app.llm_status_var.get()
        original=gui.LLMConfigDialog
        def close_dialog(parent,current):
            dialog=original(parent,current)
            assert dialog.enabled_var.get() is True
            assert dialog.base_url_var.get()=='http://example.test:8000/v1'
            assert dialog.model_var.get()=='用户指定模型'
            assert dialog.api_key_var.get()=='仅用于测试的密钥'
            dialog.enabled_var.set(False)
            dialog.after(100,dialog._save)
            return dialog
        monkeypatch.setattr(gui,'LLMConfigDialog',close_dialog)
        app.open_llm_config_dialog()
        assert app.llm_config.enabled is False
        assert '关闭' in app.llm_status_var.get()
        assert path.read_bytes()==before
    finally:
        app.destroy()
    reopened=gui.ReconciliationApp()
    try:
        reopened.withdraw()
        assert reopened.llm_config.enabled and reopened.llm_config.model=='用户指定模型'
    finally:
        reopened.destroy()


def test_没有本机配置仍可关闭启动(tmp_path,monkeypatch):
    monkeypatch.setenv('LOCALAPPDATA',str(tmp_path))
    app=gui.ReconciliationApp()
    try:
        app.withdraw()
        assert not app.llm_config.enabled and '关闭' in app.llm_status_var.get()
    finally:
        app.destroy()


@pytest.mark.parametrize('content',['{错误JSON:私人密钥}',json.dumps(['私人密钥']),json.dumps({'enabled':True,'base_url':'http://example.test/v1','model':'模型','api_key':'私人密钥','timeout_seconds':601})])
def test_配置损坏时启动提示且不泄露内容(tmp_path,monkeypatch,content):
    path=_配置(tmp_path,monkeypatch,{})
    path.write_text(content,encoding='utf-8-sig')
    messages=[]
    monkeypatch.setattr(gui.ReconciliationApp,'log',lambda self,msg:messages.append(msg))
    app=gui.ReconciliationApp()
    try:
        app.withdraw()
        assert not app.llm_config.enabled
        assert any('默认配置' in msg for msg in messages)
        assert all('私人密钥' not in msg and content not in msg for msg in messages)
    finally:
        app.destroy()


@pytest.mark.parametrize('has_local',[False,True])
def test_EXE内置配置无需本机文件且优先于本机配置(tmp_path,monkeypatch,has_local):
    import sys
    monkeypatch.setenv('LOCALAPPDATA',str(tmp_path/'local'))
    if has_local:
        _配置(tmp_path/'local',monkeypatch,{'enabled':False,'model':'旧本机模型'})
    bundle=tmp_path/'bundle'
    bundle.mkdir()
    (bundle/'大模型默认配置.json').write_text(json.dumps({'enabled':True,'mode':'online','protocol':'chat_completions','base_url':'http://example.test/v1','model':'内置模型','api_key':'内置测试密钥','timeout_seconds':600}),encoding='utf-8-sig')
    monkeypatch.setattr(gui,'__file__',str(bundle/'gui.py'))
    monkeypatch.setattr(sys,'frozen',True,raising=False)
    app=gui.ReconciliationApp()
    try:
        app.withdraw()
        assert app.llm_config.enabled
        assert app.llm_config.model=='内置模型'
        assert app.llm_config.api_key=='内置测试密钥'
        assert app._collect_run_state()['llm_config']==app.llm_config
        assert '内置模型' in app.llm_status_var.get()
    finally:
        app.destroy()


def test_源码运行不读取旁边同名包内配置(tmp_path,monkeypatch):
    monkeypatch.setenv('LOCALAPPDATA',str(tmp_path/'local'))
    (tmp_path/'大模型默认配置.json').write_text(json.dumps({'enabled':True,'base_url':'http://example.test/v1','model':'不应加载'}),encoding='utf-8-sig')
    monkeypatch.setattr(gui,'__file__',str(tmp_path/'gui.py'))
    app=gui.ReconciliationApp()
    try:
        app.withdraw()
        assert not app.llm_config.enabled
    finally:
        app.destroy()
