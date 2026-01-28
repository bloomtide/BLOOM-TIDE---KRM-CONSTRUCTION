import React, { useState, useEffect } from 'react'
import { useNavigate } from 'react-router-dom'
import { useAuth } from '../context/AuthContext'
import Logo from '../utils/logo/Logo.svg'
import LoginBackground from '../utils/logo/login_background.png'

const Login = () => {
  const navigate = useNavigate()
  const { login, isAuthenticated, loading: authLoading } = useAuth()

  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [rememberMe, setRememberMe] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');

  // Redirect if already authenticated
  useEffect(() => {
    if (isAuthenticated && !authLoading) {
      navigate('/proposals');
    }
  }, [isAuthenticated, authLoading, navigate]);

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError('');
    setLoading(true);

    try {
      const result = await login(email, password);

      if (result.success) {
        // Redirect to proposals page
        navigate('/proposals');
      } else {
        setError(result.message || 'Login failed');
      }
    } catch (err) {
      setError('An error occurred. Please try again.');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="flex h-screen">
      {/* Left Side - Blue Background with KRM Branding */}
      <div className="hidden lg:flex lg:w-1/2 relative overflow-hidden" style={{ backgroundImage: `url(${LoginBackground})`, backgroundSize: 'cover', backgroundPosition: 'center' }}>
        {/* Background overlay for better text readability */}
        <div className="absolute inset-0 bg-[#2B4A8C]/40"></div>

        <div className="relative z-10 flex flex-col justify-center px-16 text-white">
          {/* KRM Logo */}
          <div className="mb-12">
            <img src={Logo} alt="KRM Construction" className="h-16 w-auto" />
          </div>

          {/* Main Heading */}
          <h1 className="text-4xl font-bold mb-6 leading-tight">
            Secure access to KRM<br />systems
          </h1>

          {/* Subtext */}
          <p className="text-lg opacity-90 leading-relaxed max-w-md">
            Log in to access authorized construction workflows,<br />
            project data, and internal tools.
          </p>
        </div>
      </div>

      {/* Right Side - White Background with Login Form */}
      <div className="flex-1 flex items-center justify-center bg-gray-50 px-8 lg:px-16">
        <div className="w-full max-w-md">
          {/* Welcome Text */}
          <div className="mb-10 text-center">
            <h2 className="text-3xl font-bold text-gray-900 mb-3">Welcome back!</h2>
            <p className="text-sm text-gray-500">
              Log in to manage your projects, calculations, and proposals in<br />one place.
            </p>
          </div>

          {/* Error Message */}
          {error && (
            <div className="mb-6 p-4 bg-red-50 border border-red-200 rounded-lg">
              <p className="text-sm text-red-600">{error}</p>
            </div>
          )}

          {/* Login Form */}
          <form onSubmit={handleSubmit} className="space-y-6">
            {/* Email Field */}
            <div>
              <label htmlFor="email" className="block text-sm font-medium text-gray-700 mb-2">
                Email
              </label>
              <input
                type="email"
                id="email"
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                placeholder="Enter your email"
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent outline-none transition-all text-gray-900 placeholder:text-gray-400"
                required
              />
            </div>

            {/* Password Field */}
            <div>
              <label htmlFor="password" className="block text-sm font-medium text-gray-700 mb-2">
                Password
              </label>
              <div className="relative">
                <input
                  type={showPassword ? "text" : "password"}
                  id="password"
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  placeholder="Enter your password"
                  className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent outline-none transition-all text-gray-900 placeholder:text-gray-400 pr-12"
                  required
                />
                <button
                  type="button"
                  onClick={() => setShowPassword(!showPassword)}
                  className="absolute right-4 top-1/2 -translate-y-1/2 text-gray-400 hover:text-gray-600 transition-colors"
                >
                  {showPassword ? (
                    <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M13.875 18.825A10.05 10.05 0 0112 19c-4.478 0-8.268-2.943-9.543-7a9.97 9.97 0 011.563-3.029m5.858.908a3 3 0 114.243 4.243M9.878 9.878l4.242 4.242M9.88 9.88l-3.29-3.29m7.532 7.532l3.29 3.29M3 3l3.59 3.59m0 0A9.953 9.953 0 0112 5c4.478 0 8.268 2.943 9.543 7a10.025 10.025 0 01-4.132 5.411m0 0L21 21" />
                    </svg>
                  ) : (
                    <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M15 12a3 3 0 11-6 0 3 3 0 016 0z" />
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M2.458 12C3.732 7.943 7.523 5 12 5c4.478 0 8.268 2.943 9.542 7-1.274 4.057-5.064 7-9.542 7-4.477 0-8.268-2.943-9.542-7z" />
                    </svg>
                  )}
                </button>
              </div>
            </div>

            {/* Remember Me Checkbox */}
            <div className="flex items-center">
              <input
                type="checkbox"
                id="remember"
                checked={rememberMe}
                onChange={(e) => setRememberMe(e.target.checked)}
                className="w-4 h-4 text-blue-600 border-gray-300 rounded focus:ring-2 focus:ring-blue-500"
              />
              <label htmlFor="remember" className="ml-2 text-sm text-gray-700">
                Remember me
              </label>
            </div>

            {/* Login Button */}
            <button
              type="submit"
              disabled={loading}
              className="w-full bg-[#1A72B9] hover:bg-[#1565C0] disabled:bg-gray-400 disabled:cursor-not-allowed text-white text-sm font-medium py-3 px-4 rounded-lg transition-colors duration-200 shadow-sm flex items-center justify-center gap-2"
              style={{ fontFamily: 'Plus Jakarta Sans' }}
            >
              {loading && (
                <svg className="animate-spin h-5 w-5" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
                  <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle>
                  <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
                </svg>
              )}
              {loading ? 'Logging in...' : 'Log In'}
            </button>
          </form>

          {/* Footer Link */}
          <div className="mt-8 text-center text-sm text-gray-600">
            Don't have an account?{' '}
            <a href="#" className="text-[#1976D2] hover:text-[#1565C0] font-medium transition-colors">
              Contact Admin
            </a>
          </div>
        </div>
      </div>
    </div>
  )
}

export default Login